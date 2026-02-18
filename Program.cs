using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace Ak0Analyzer
{
    public class MainForm : Form
    {
        private CheckedListBox clbWarehouses;
        private Button btnRun, btnSelectFolder, btnLoadSchedule, btnSettings;
        private CheckBox chkEnableUPS, chkFilterI, chkFilterE;
        private Label lblStatus;
        private List<FileItem> sortedFiles;
        private HashSet<string> allDetectedLocs = new HashSet<string>();
        private string selectedFolderPath = "";
        private Dictionary<ScheduleKey, string> staffSchedule = new Dictionary<ScheduleKey, string>();

        private string upsLicense = "", upsUser = "", upsPass = "";
        private readonly string settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ups_settings.ini");

        struct FileItem { public string Path; public DateTime Date; }
        struct ScheduleKey { 
            public string Loc; public int Day;
            public override bool Equals(object obj) => obj is ScheduleKey other && Loc == other.Loc && Day == other.Day;
            public override int GetHashCode() => (Loc?.GetHashCode() ?? 0) ^ Day.GetHashCode();
        }

        public MainForm()
        {
            LoadSettings();
            this.Text = "WSQA PRO - AK0 Analyzer";
            this.Size = new System.Drawing.Size(550, 850);
            this.StartPosition = FormStartPosition.CenterScreen;

            FlowLayoutPanel topPanel = new FlowLayoutPanel() { Dock = DockStyle.Top, Height = 220, Padding = new Padding(10) };
            
            btnSelectFolder = new Button() { Text = "📁 1. WYBIERZ FOLDER AK0", Size = new System.Drawing.Size(245, 60), BackColor = System.Drawing.Color.LightSkyBlue, FlatStyle = FlatStyle.Flat, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
            btnSelectFolder.Click += (s, e) => SelectFolder();
            
            btnLoadSchedule = new Button() { Text = "📅 2. WCZYTAJ GRAFIK", Size = new System.Drawing.Size(245, 60), BackColor = System.Drawing.Color.NavajoWhite, FlatStyle = FlatStyle.Flat, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
            btnLoadSchedule.Click += (s, e) => LoadScheduleWindow();
            
            btnSettings = new Button() { Text = "⚙️ USTAWIENIA UPS API", Size = new System.Drawing.Size(500, 40), BackColor = System.Drawing.Color.LightGray, FlatStyle = FlatStyle.Flat };
            btnSettings.Click += (s, e) => ShowSettingsWindow();

            // Filtry Import/Export
            GroupBox gpFilters = new GroupBox() { Text = "Filtry magazynów", Size = new System.Drawing.Size(500, 50) };
            chkFilterI = new CheckBox() { Text = "Import (I...)", Checked = true, AutoSize = true, Location = new System.Drawing.Point(10, 20) };
            chkFilterE = new CheckBox() { Text = "Export (E...)", Checked = true, AutoSize = true, Location = new System.Drawing.Point(150, 20) };
            chkFilterI.CheckedChanged += (s, e) => ApplyLocFilter();
            chkFilterE.CheckedChanged += (s, e) => ApplyLocFilter();
            gpFilters.Controls.Add(chkFilterI); gpFilters.Controls.Add(chkFilterE);

            topPanel.Controls.Add(btnSelectFolder);
            topPanel.Controls.Add(btnLoadSchedule);
            topPanel.Controls.Add(btnSettings);
            topPanel.Controls.Add(gpFilters);

            clbWarehouses = new CheckedListBox() { Dock = DockStyle.Fill, CheckOnClick = true, Font = new System.Drawing.Font("Segoe UI", 10) };
            
            Panel pnlOptions = new Panel() { Dock = DockStyle.Bottom, Height = 40, BackColor = System.Drawing.Color.WhiteSmoke };
            chkEnableUPS = new CheckBox() { Text = "Automatyczna weryfikacja UPS (Auto-Green)", AutoSize = true, Location = new System.Drawing.Point(10, 10), Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
            pnlOptions.Controls.Add(chkEnableUPS);

            btnRun = new Button() { Text = "🚀 3. GENERUJ RAPORT", Dock = DockStyle.Bottom, Height = 70, BackColor = System.Drawing.Color.LightGreen, Enabled = false, Font = new System.Drawing.Font("Segoe UI", 11, System.Drawing.FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btnRun.Click += BtnRun_Click;

            lblStatus = new Label() { Text = "Gotowy", Dock = DockStyle.Bottom, Height = 40, TextAlign = System.Drawing.ContentAlignment.MiddleCenter, BackColor = System.Drawing.Color.WhiteSmoke, BorderStyle = BorderStyle.FixedSingle };

            this.Controls.Add(clbWarehouses);
            this.Controls.Add(new Label() { Text = " Magazyny do analizy:", Dock = DockStyle.Top, Height = 25, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) });
            this.Controls.Add(topPanel);
            this.Controls.Add(pnlOptions);
            this.Controls.Add(lblStatus);
            this.Controls.Add(btnRun);
        }

        private void ApplyLocFilter() {
            clbWarehouses.Items.Clear();
            foreach (var loc in allDetectedLocs.OrderBy(x => x)) {
                bool isI = loc.StartsWith("I", StringComparison.OrdinalIgnoreCase);
                bool isE = loc.StartsWith("E", StringComparison.OrdinalIgnoreCase);
                if ((isI && chkFilterI.Checked) || (isE && chkFilterE.Checked)) {
                    clbWarehouses.Items.Add(loc);
                }
            }
        }

        private void SelectFolder() {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog()) {
                if (fbd.ShowDialog() == DialogResult.OK) {
                    selectedFolderPath = fbd.SelectedPath;
                    ScanFiles();
                }
            }
        }

        private void ScanFiles() {
            allDetectedLocs.Clear();
            var files = Directory.GetFiles(selectedFolderPath, "*.xlsx");
            var valid = new List<FileItem>();
            foreach (var f in files) {
                string fn = Path.GetFileName(f);
                var m = Regex.Match(fn, @"(\d{2}\.\d{2}\.\d{4})");
                if (m.Success && fn.ToUpper().StartsWith("AK0"))
                    if (DateTime.TryParseExact(m.Value, "dd.MM.yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime dt))
                        valid.Add(new FileItem { Path = f, Date = dt });
            }
            sortedFiles = valid.OrderBy(x => x.Date).ToList();
            if (sortedFiles.Count < 2) { lblStatus.Text = "Błąd: Potrzeba min. 2 plików!"; return; }

            foreach (var f in sortedFiles) {
                using (var wb = new XLWorkbook(f.Path)) {
                    var ws = wb.Worksheets.FirstOrDefault(w => w.Name.ToUpper().Contains("AK0")) ?? wb.Worksheets.FirstOrDefault();
                    if (ws == null) continue;
                    var range = ws.RangeUsed(); if (range == null) continue;
                    foreach (var row in range.RowsUsed().Skip(1)) {
                        string val = row.Cell(1).GetString().Trim();
                        if (!string.IsNullOrEmpty(val)) allDetectedLocs.Add(val);
                    }
                }
            }
            ApplyLocFilter();
            btnRun.Enabled = true;
            lblStatus.Text = "Wczytano " + sortedFiles.Count + " plików.";
        }

        private void LoadScheduleWindow() {
            Form f = new Form() { Text = "Wklej Grafik (A1 = Miesiąc, A2.. = Mag)", Size = new System.Drawing.Size(800, 500), StartPosition = FormStartPosition.CenterParent };
            DataGridView dgv = new DataGridView() { Dock = DockStyle.Fill, AllowUserToAddRows = false };
            Button btnSave = new Button() { Text = "Zapisz i Mapuj Grafik", Dock = DockStyle.Bottom, Height = 45, BackColor = System.Drawing.Color.PaleGreen };
            dgv.KeyDown += (s, e) => { if (e.Control && e.KeyCode == Keys.V) PasteToDgv(dgv); };
            btnSave.Click += (s, e) => { ProcessSchedule(dgv); f.Close(); };
            f.Controls.Add(dgv); f.Controls.Add(btnSave); f.ShowDialog();
        }

        private void PasteToDgv(DataGridView dgv) {
            string t = Clipboard.GetText(); if (string.IsNullOrEmpty(t)) return;
            dgv.Rows.Clear(); dgv.Columns.Clear();
            string[] lines = t.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0) return;
            string[] headers = lines[0].Split('\t');
            foreach (var h in headers) dgv.Columns.Add(h, h);
            for (int i = 1; i < lines.Length; i++) dgv.Rows.Add(lines[i].Split('\t'));
        }

        private void ProcessSchedule(DataGridView dgv) {
            staffSchedule.Clear();
            for (int r = 0; r < dgv.Rows.Count; r++) {
                string rawLoc = dgv.Rows[r].Cells[0].Value?.ToString().Trim().ToLower() ?? "";
                if (string.IsNullOrEmpty(rawLoc)) continue;

                // Tłumaczenie nazw: Mag1 -> IWMAGAZYN1, Smalls -> IWMSMALLS
                List<string> mappedLocs = new List<string>();
                if (rawLoc.Contains("mag")) {
                    string num = Regex.Match(rawLoc, @"\d+").Value;
                    mappedLocs.Add("IWMAGAZYN" + num);
                    mappedLocs.Add("EWMAGEXP" + num); // Na wypadek Exportu
                } else if (rawLoc.Contains("smalls")) {
                    mappedLocs.Add("IWMSMALLSXX");
                    mappedLocs.Add("IWMSMALLS1");
                }

                for (int c = 1; c < dgv.Columns.Count; c++) {
                    string dayHeader = Regex.Match(dgv.Columns[c].HeaderText, @"\d+").Value;
                    if (int.TryParse(dayHeader, out int day)) {
                        string person = dgv.Rows[r].Cells[c].Value?.ToString().Trim() ?? "";
                        if (!string.IsNullOrEmpty(person)) {
                            // Zamiana nowych linii na separator jeśli są 2 osoby
                            person = person.Replace("\r\n", " / ").Replace("\n", " / ");
                            foreach (var ml in mappedLocs) {
                                staffSchedule[new ScheduleKey { Loc = ml.ToUpper(), Day = day }] = person;
                            }
                        }
                    }
                }
            }
            MessageBox.Show("Zmapowano grafik dla " + staffSchedule.Keys.Select(x => x.Loc).Distinct().Count() + " magazynów.");
        }

        private async void BtnRun_Click(object sender, EventArgs e) {
            btnRun.Enabled = false;
            try { await GenerateReportAsync(); MessageBox.Show("Raport gotowy!"); }
            catch (Exception ex) { MessageBox.Show("Błąd: " + ex.Message); }
            finally { btnRun.Enabled = true; lblStatus.Text = "Gotowe."; }
        }

        private async System.Threading.Tasks.Task GenerateReportAsync() {
            var selectedLocs = clbWarehouses.CheckedItems.Cast<string>().ToList();
            var data = new Dictionary<string, SortedDictionary<DateTime, string>>();
            var dates = sortedFiles.Select(x => x.Date).ToList();
            DateTime lastDay = dates.Max();
            var personStats = new Dictionary<string, int>();

            foreach (var f in sortedFiles) {
                using (var wb = new XLWorkbook(f.Path)) {
                    var ws = wb.Worksheets.FirstOrDefault(w => w.Name.ToUpper().Contains("AK0")) ?? wb.Worksheets.FirstOrDefault();
                    var range = ws?.RangeUsed(); if (range == null) continue;
                    foreach (var row in range.RowsUsed().Skip(1)) {
                        string l = row.Cell(1).GetString().Trim();
                        string p = row.Cell(2).GetString().Trim();
                        if (selectedLocs.Contains(l)) {
                            if (!data.ContainsKey(p)) data[p] = new SortedDictionary<DateTime, string>();
                            data[p][f.Date] = l;
                        }
                    }
                }
            }

            using (var report = new XLWorkbook()) {
                var ws = report.Worksheets.Add("Analiza");
                ws.Cell(1, 1).Value = "Package ID";
                for (int i = 0; i < dates.Count; i++) ws.Cell(1, i + 2).Value = dates[i].ToShortDateString();

                int r = 2;
                foreach (var pkg in data) {
                    DateTime first = pkg.Value.Keys.Min();
                    DateTime lastScan = pkg.Value.Keys.Max();
                    bool missing = !pkg.Value.ContainsKey(lastDay) || pkg.Value.Count < dates.Count(d => d >= first && d <= lastScan);

                    if (missing) {
                        ws.Cell(r, 1).Value = pkg.Key;
                        for (int i = 0; i < dates.Count; i++) {
                            DateTime d = dates[i];
                            if (pkg.Value.ContainsKey(d)) ws.Cell(r, i + 2).Value = pkg.Value[d];
                            else if (d > first) {
                                var cell = ws.Cell(r, i + 2);
                                cell.Value = "BRAK SKANU"; cell.Style.Fill.BackgroundColor = XLColor.Salmon;
                                
                                // Próba dopasowania osoby
                                string currentLoc = pkg.Value[first].ToUpper();
                                var key = new ScheduleKey { Loc = currentLoc, Day = d.Day };
                                if (staffSchedule.TryGetValue(key, out string p)) {
                                    cell.CreateComment().AddText(p);
                                    if (!personStats.ContainsKey(p)) personStats[p] = 0;
                                    personStats[p]++;
                                }
                            }
                        }
                        r++;
                    }
                }
                ws.Columns().AdjustToContents();
                report.SaveAs(Path.Combine(selectedFolderPath, "Raport_Brakow_AK0_" + DateTime.Now.ToString("ddMMyy_HHmm") + ".xlsx"));
            }
        }

        private void LoadSettings() { if (File.Exists(settingsPath)) { var lines = File.ReadAllLines(settingsPath); if (lines.Length >= 3) { upsLicense = lines[0]; upsUser = lines[1]; upsPass = lines[2]; } } }
        private void ShowSettingsWindow() { /* ... identyczna jak wcześniej ... */ }

        [STAThread] static void Main() { Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false); Application.Run(new MainForm()); }
    }
}
