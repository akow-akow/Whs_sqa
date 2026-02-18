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
        private CheckBox chkEnableUPS;
        private Label lblStatus;
        private List<(string Path, DateTime Date)> sortedFiles;
        private string selectedFolderPath = "";
        private Dictionary<(string Loc, int Day), string> staffSchedule = new Dictionary<(string, int), string>();

        // Ustawienia UPS
        private string upsLicense = "", upsUser = "", upsPass = "";
        private readonly string settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ups_settings.ini");

        // Statystyki API
        private int apiSuccess = 0;
        private int apiFailed = 0;

        public MainForm()
        {
            LoadSettings();
            this.Text = "WSQA PRO + UPS Tracking";
            this.Size = new System.Drawing.Size(550, 830);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Ustawienie ikony okna z pliku icon.ico (zgodnie z Twoim csproj)
            try {
                string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "icon.ico");
                if (File.Exists(iconPath)) this.Icon = new System.Drawing.Icon(iconPath);
            } catch { }

            // Panel górny - Przyciski akcji
            FlowLayoutPanel topPanel = new FlowLayoutPanel() { Dock = DockStyle.Top, Height = 180, Padding = new Padding(10) };
            
            btnSelectFolder = new Button() { Text = "📁 1. WYBIERZ FOLDER AK0", Size = new System.Drawing.Size(245, 60), BackColor = System.Drawing.Color.LightSkyBlue, FlatStyle = FlatStyle.Flat, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
            btnSelectFolder.Click += (s, e) => SelectFolder();

            btnLoadSchedule = new Button() { Text = "📅 2. WCZYTAJ GRAFIK", Size = new System.Drawing.Size(245, 60), BackColor = System.Drawing.Color.NavajoWhite, FlatStyle = FlatStyle.Flat, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
            btnLoadSchedule.Click += (s, e) => LoadScheduleWindow();

            btnSettings = new Button() { Text = "⚙️ USTAWIENIA UPS API", Size = new System.Drawing.Size(500, 40), BackColor = System.Drawing.Color.LightGray, FlatStyle = FlatStyle.Flat };
            btnSettings.Click += (s, e) => ShowSettingsWindow();

            topPanel.Controls.Add(btnSelectFolder);
            topPanel.Controls.Add(btnLoadSchedule);
            topPanel.Controls.Add(btnSettings);

            // Lista magazynów
            clbWarehouses = new CheckedListBox() { Dock = DockStyle.Fill, CheckOnClick = true, Font = new System.Drawing.Font("Segoe UI", 10) };
            
            // Panel opcji UPS
            Panel pnlOptions = new Panel() { Dock = DockStyle.Bottom, Height = 40, BackColor = System.Drawing.Color.WhiteSmoke };
            chkEnableUPS = new CheckBox() { Text = "Weryfikuj braki przez UPS API (ostatni dzień)", AutoSize = true, Location = new System.Drawing.Point(10, 10), Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
            pnlOptions.Controls.Add(chkEnableUPS);

            // Przycisk generowania
            btnRun = new Button() { Text = "🚀 3. GENERUJ RAPORT", Dock = DockStyle.Bottom, Height = 70, BackColor = System.Drawing.Color.LightGreen, Enabled = false, Font = new System.Drawing.Font("Segoe UI", 11, System.Drawing.FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btnRun.Click += BtnRun_Click;

            // Status bar
            lblStatus = new Label() { Text = "Gotowy", Dock = DockStyle.Bottom, Height = 40, TextAlign = System.Drawing.ContentAlignment.MiddleCenter, BackColor = System.Drawing.Color.WhiteSmoke, BorderStyle = BorderStyle.FixedSingle };

            this.Controls.Add(clbWarehouses);
            this.Controls.Add(new Label() { Text = " Magazyny do analizy:", Dock = DockStyle.Top, Height = 25, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) });
            this.Controls.Add(topPanel);
            this.Controls.Add(pnlOptions);
            this.Controls.Add(lblStatus);
            this.Controls.Add(btnRun);
        }

        private void LoadSettings()
        {
            if (File.Exists(settingsPath)) {
                var lines = File.ReadAllLines(settingsPath);
                if (lines.Length >= 3) { upsLicense = lines[0]; upsUser = lines[1]; upsPass = lines[2]; }
            }
        }

        private void ShowSettingsWindow()
        {
            Form f = new Form() { Text = "Konfiguracja UPS API", Size = new System.Drawing.Size(400, 250), StartPosition = FormStartPosition.CenterParent };
            TableLayoutPanel tlp = new TableLayoutPanel() { Dock = DockStyle.Fill, Padding = new Padding(10), ColumnCount = 2 };
            tlp.Controls.Add(new Label() { Text = "License Number:" }, 0, 0);
            TextBox txtL = new TextBox() { Text = upsLicense, Width = 200 }; tlp.Controls.Add(txtL, 1, 0);
            tlp.Controls.Add(new Label() { Text = "User ID:" }, 0, 1);
            TextBox txtU = new TextBox() { Text = upsUser, Width = 200 }; tlp.Controls.Add(txtU, 1, 1);
            tlp.Controls.Add(new Label() { Text = "Password:" }, 0, 2);
            TextBox txtP = new TextBox() { Text = upsPass, Width = 200, UseSystemPasswordChar = true }; tlp.Controls.Add(txtP, 1, 2);
            Button btnS = new Button() { Text = "ZAPISZ", Dock = DockStyle.Bottom, Height = 40 };
            btnS.Click += (s, e) => {
                upsLicense = txtL.Text; upsUser = txtU.Text; upsPass = txtP.Text;
                File.WriteAllLines(settingsPath, new[] { upsLicense, upsUser, upsPass });
                f.Close();
            };
            f.Controls.Add(tlp); f.Controls.Add(btnS); f.ShowDialog();
        }

        private void LoadScheduleWindow()
        {
            Form f = new Form() { Text = "Arkusz Grafiku - Ctrl+V aby wkleić", Size = new System.Drawing.Size(800, 500), StartPosition = FormStartPosition.CenterParent };
            DataGridView dgv = new DataGridView() { Dock = DockStyle.Fill, BackgroundColor = System.Drawing.Color.White, AllowUserToAddRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells };
            
            Panel pnl = new Panel() { Dock = DockStyle.Bottom, Height = 50 };
            Button btnClear = new Button() { Text = "WYCZYŚĆ TABELĘ", Dock = DockStyle.Left, Width = 150, BackColor = System.Drawing.Color.MistyRose };
            Button btnSave = new Button() { Text = "ZAPISZ I ZAMKNIJ", Dock = DockStyle.Fill, BackColor = System.Drawing.Color.PaleGreen, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };

            dgv.KeyDown += (s, e) => { if (e.Control && e.KeyCode == Keys.V) PasteToDgv(dgv); };
            btnClear.Click += (s, e) => { dgv.Rows.Clear(); dgv.Columns.Clear(); };
            btnSave.Click += (s, e) => { ProcessDgvData(dgv); f.Close(); };

            pnl.Controls.Add(btnSave); pnl.Controls.Add(btnClear);
            f.Controls.Add(dgv); f.Controls.Add(pnl); f.ShowDialog();
        }

        private void PasteToDgv(DataGridView dgv)
        {
            string text = Clipboard.GetText();
            if (string.IsNullOrEmpty(text)) return;
            dgv.Rows.Clear(); dgv.Columns.Clear();
            string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0) return;
            
            string[] headers = lines[0].Split('\t');
            foreach (var h in headers) dgv.Columns.Add(h, h);
            for (int i = 1; i < lines.Length; i++) dgv.Rows.Add(lines[i].Split('\t'));
        }

        private void ProcessDgvData(DataGridView dgv)
        {
            staffSchedule.Clear();
            for (int r = 0; r < dgv.Rows.Count; r++) {
                string rawLoc = dgv.Rows[r].Cells[0].Value?.ToString() ?? "";
                string mappedLoc = MapLocationName(rawLoc.ToLower().Trim());
                for (int c = 1; c < dgv.Columns.Count; c++) {
                    if (int.TryParse(dgv.Columns[c].HeaderText, out int dayNum)) {
                        string person = dgv.Rows[r].Cells[c].Value?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(person)) staffSchedule[(mappedLoc, dayNum)] = person;
                    }
                }
            }
            lblStatus.Text = "Grafik przetworzony.";
        }

        private string MapLocationName(string raw)
        {
            if (raw.Contains("100")) return "IWMAG100";
            if (raw.Contains("mag")) {
                var m = Regex.Match(raw, @"\d+");
                return m.Success ? "IWMAGAZYN" + m.Value : raw.ToUpper();
            }
            if (raw.Contains("smalls")) return "IWMSMALLS1";
            return raw.ToUpper();
        }

        private void SelectFolder()
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog() { AutoUpgradeEnabled = false }) {
                if (fbd.ShowDialog() == DialogResult.OK) { selectedFolderPath = fbd.SelectedPath; ScanFiles(); }
            }
        }

        private void ScanFiles()
        {
            clbWarehouses.Items.Clear();
            btnRun.Enabled = false;
            if (string.IsNullOrEmpty(selectedFolderPath) || !Directory.Exists(selectedFolderPath)) return;

            var files = Directory.GetFiles(selectedFolderPath, "*.xlsx");
            var valid = new List<(string, DateTime)>();
            foreach (var f in files) {
                string fn = Path.GetFileName(f);
                if (fn.StartsWith("~$")) continue;
                var m = Regex.Match(fn, @"(\d{2}\.\d{2}\.\d{4})");
                if (m.Success && fn.ToUpper().StartsWith("AK0"))
                    if (DateTime.TryParseExact(m.Value, "dd.MM.yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime dt))
                        valid.Add((f, dt));
            }
            sortedFiles = valid.OrderBy(x => x.Item2).ToList();
            if (sortedFiles.Count < 2) { lblStatus.Text = "Błąd: Min. 2 pliki AK0!"; return; }

            HashSet<string> locs = new HashSet<string>();
            try {
                foreach (var f in sortedFiles) {
                    using (var wb = new XLWorkbook(f.Item1)) {
                        var ws = wb.Worksheets.FirstOrDefault(w => w.Name.ToUpper() == "AK0") ?? wb.Worksheets.FirstOrDefault();
                        if (ws == null) continue;
                        var used = ws.RangeUsed(); 
                        if (used == null) continue; // FIX BŁĘDU JIT
                        foreach (var row in used.RowsUsed().Skip(1)) {
                            string l = row.Cell(1).GetString().Trim();
                            if (l.StartsWith("I", StringComparison.OrdinalIgnoreCase)) locs.Add(l);
                        }
                    }
                }
                foreach (var l in locs.OrderBy(x => x)) clbWarehouses.Items.Add(l, false);
                btnRun.Enabled = true;
                lblStatus.Text = $"Wczytano {sortedFiles.Count} dni.";
            } catch (Exception ex) { MessageBox.Show("Błąd skanowania: " + ex.Message); }
        }

        private async void BtnRun_Click(object sender, EventArgs e)
        {
            if (clbWarehouses.CheckedItems.Count == 0) { MessageBox.Show("Proszę wybrać magazyny do analizy!"); return; }
            btnRun.Enabled = false;
            apiSuccess = 0; apiFailed = 0;
            lblStatus.Text = "Generowanie raportu...";
            Application.DoEvents();
            try { 
                await GenerateReportAsync(); 
                lblStatus.Text = "Raport gotowy!";
                MessageBox.Show("Raport został zapisany w folderze AK0."); 
            }
            catch (Exception ex) { MessageBox.Show("Błąd: " + ex.Message); }
            finally { btnRun.Enabled = true; }
        }

        private async System.Threading.Tasks.Task GenerateReportAsync()
        {
            var selectedLocs = clbWarehouses.CheckedItems.Cast<string>().ToList();
            var data = new Dictionary<string, SortedDictionary<DateTime, string>>();
            var dates = sortedFiles.Select(x => x.Item2).ToList();
            DateTime lastDay = dates.Max();

            // Odczyt danych z plików AK0
            foreach (var f in sortedFiles) {
                using (var wb = new XLWorkbook(f.Item1)) {
                    var ws = wb.Worksheets.FirstOrDefault(w => w.Name.ToUpper() == "AK0") ?? wb.Worksheets.FirstOrDefault();
                    if (ws == null) continue;
                    var used = ws.RangeUsed(); if (used == null) continue;
                    foreach (var row in used.RowsUsed().Skip(1)) {
                        string l = row.Cell(1).GetString().Trim();
                        string p = row.Cell(2).GetString().Trim();
                        if (selectedLocs.Contains(l)) {
                            if (!data.ContainsKey(p)) data[p] = new SortedDictionary<DateTime, string>();
                            data[p][f.Item2] = l;
                        }
                    }
                }
            }

            using (var report = new XLWorkbook()) {
                var ws = report.Worksheets.Add("Analiza");
                var wsStat = report.Worksheets.Add("Statystyki");
                
                ws.Cell(1, 1).Value = "Package ID";
                for (int i = 0; i < dates.Count; i++) ws.Cell(1, i + 2).Value = dates[i].ToShortDateString();
                int colUPS = dates.Count + 2;
                ws.Cell(1, colUPS).Value = "Ostatni Status UPS";

                int r = 2;
                foreach (var pkg in data) {
                    var first = pkg.Value.Keys.Min();
                    var lastScan = pkg.Value.Keys.Max();
                    
                    // Luka wewnątrz okresu
                    bool hasGap = false;
                    for (int i = 0; i < dates.Count; i++) {
                        if (dates[i] > first && dates[i] < lastScan && !pkg.Value.ContainsKey(dates[i])) {
                            var next = pkg.Value.Keys.Where(d => d > dates[i]).Min();
                            if ((next - dates[i]).TotalDays <= 3) { hasGap = true; break; }
                        }
                    }
                    // Brak w ostatnim dniu
                    bool missingLast = !pkg.Value.ContainsKey(lastDay) && (lastDay - lastScan).TotalDays <= 3;

                    if (hasGap || missingLast) {
                        ws.Cell(r, 1).Value = pkg.Key;
                        for (int i = 0; i < dates.Count; i++) {
                            DateTime d = dates[i];
                            if (pkg.Value.ContainsKey(d)) ws.Cell(r, i + 2).Value = pkg.Value[d];
                            else if (d > first) {
                                var cell = ws.Cell(r, i + 2);
                                cell.Value = "BRAK SKANU"; cell.Style.Fill.BackgroundColor = XLColor.Salmon;
                                if (staffSchedule.TryGetValue((pkg.Value[first], d.Day), out string person)) cell.CreateComment().AddText(person);
                            }
                        }
                        // API UPS
                        if (missingLast && chkEnableUPS.Checked && !string.IsNullOrEmpty(upsLicense)) {
                            lblStatus.Text = $"UPS Tracking: {pkg.Key}..."; 
                            Application.DoEvents();
                            ws.Cell(r, colUPS).Value = await GetUpsTracking(pkg.Key);
                        }
                        r++;
                    }
                }

                // Wypełnianie statystyk
                wsStat.Cell(1, 1).Value = "RAPORT STATYSTYCZNY";
                wsStat.Cell(1, 1).Style.Font.Bold = true;
                wsStat.Cell(3, 1).Value = "Pomyślne zapytania UPS:"; wsStat.Cell(3, 2).Value = apiSuccess;
                wsStat.Cell(4, 1).Value = "Błędy/Brak danych UPS:"; wsStat.Cell(4, 2).Value = apiFailed;
                wsStat.Cell(5, 1).Value = "Łącznie sprawdzonych:"; wsStat.Cell(5, 2).Value = apiSuccess + apiFailed;
                wsStat.Cell(7, 1).Value = "Data wygenerowania:"; wsStat.Cell(7, 2).Value = DateTime.Now.ToString("g");

                ws.Columns().AdjustToContents();
                wsStat.Columns().AdjustToContents();
                report.SaveAs(Path.Combine(selectedFolderPath, $"Raport_Analiza_{DateTime.Now:ddMMyy_HHmm}.xlsx"));
            }
        }

        private async System.Threading.Tasks.Task<string> GetUpsTracking(string trackNum)
        {
            try {
                string xml = $@"<?xml version=""1.0""?><AccessRequest><AccessLicenseNumber>{upsLicense}</AccessLicenseNumber><UserId>{upsUser}</UserId><Password>{upsPass}</Password></AccessRequest>" +
                             $@"<?xml version=""1.0""?><TrackRequest><Request><RequestAction>Track</RequestAction></Request><TrackingNumber>{trackNum}</TrackingNumber></TrackRequest>";
                using (var client = new HttpClient()) {
                    var resp = await client.PostAsync("https://www.ups.com/ups.app/xml/Track", new StringContent(xml, Encoding.UTF8, "application/x-www-form-urlencoded"));
                    var content = await resp.Content.ReadAsStringAsync();
                    var doc = XDocument.Parse(content);
                    var package = doc.Descendants("Package").FirstOrDefault();
                    if (package != null) {
                        var lastActivity = package.Descendants("Activity").FirstOrDefault();
                        string desc = lastActivity?.Descendants("Status")?.FirstOrDefault()?.Descendants("StatusType")?.FirstOrDefault()?.Descendants("Description")?.FirstOrDefault()?.Value ?? "Brak opisu";
                        string city = lastActivity?.Descendants("ActivityLocation")?.FirstOrDefault()?.Descendants("Address")?.FirstOrDefault()?.Descendants("City")?.FirstOrDefault()?.Value ?? "";
                        apiSuccess++;
                        return city != "" ? $"{desc} - {city}" : desc;
                    }
                }
            } catch { }
            apiFailed++;
            return "Błąd API / Brak danych";
        }

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
