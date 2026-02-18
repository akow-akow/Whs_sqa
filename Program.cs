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
        private List<FileItem> sortedFiles;
        private string selectedFolderPath = "";
        private Dictionary<ScheduleKey, string> staffSchedule = new Dictionary<ScheduleKey, string>();

        private string upsLicense = "", upsUser = "", upsPass = "";
        private readonly string settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ups_settings.ini");

        private int apiSuccess = 0;
        private int apiFailed = 0;

        struct FileItem { public string Path; public DateTime Date; }
        struct ScheduleKey { 
            public string Loc; public int Day;
            public override bool Equals(object obj) => obj is ScheduleKey other && Loc == other.Loc && Day == other.Day;
            public override int GetHashCode() => (Loc?.GetHashCode() ?? 0) ^ Day.GetHashCode();
        }

        public MainForm()
        {
            LoadSettings();
            this.Text = "WSQA PRO - Lightweight Edition";
            this.Size = new System.Drawing.Size(550, 830);
            this.StartPosition = FormStartPosition.CenterScreen;

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

            clbWarehouses = new CheckedListBox() { Dock = DockStyle.Fill, CheckOnClick = true, Font = new System.Drawing.Font("Segoe UI", 10) };
            Panel pnlOptions = new Panel() { Dock = DockStyle.Bottom, Height = 40, BackColor = System.Drawing.Color.WhiteSmoke };
            chkEnableUPS = new CheckBox() { Text = "Automatyczna weryfikacja lokalizacji (Auto-Green)", AutoSize = true, Location = new System.Drawing.Point(10, 10), Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
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

        private void LoadSettings() { if (File.Exists(settingsPath)) { var lines = File.ReadAllLines(settingsPath); if (lines.Length >= 3) { upsLicense = lines[0]; upsUser = lines[1]; upsPass = lines[2]; } } }

        private void SelectFolder() { using (FolderBrowserDialog fbd = new FolderBrowserDialog()) { if (fbd.ShowDialog() == DialogResult.OK) { selectedFolderPath = fbd.SelectedPath; ScanFiles(); } } }

        private void ScanFiles() {
            clbWarehouses.Items.Clear();
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
            
            HashSet<string> locs = new HashSet<string>();
            foreach (var f in sortedFiles) {
                using (var wb = new XLWorkbook(f.Path)) {
                    var ws = wb.Worksheets.FirstOrDefault();
                    if (ws == null) continue;
                    foreach (var row in ws.RangeUsed().RowsUsed().Skip(1)) locs.Add(row.Cell(1).GetString().Trim());
                }
            }
            foreach (var l in locs.OrderBy(x => x)) clbWarehouses.Items.Add(l);
            btnRun.Enabled = true; lblStatus.Text = "Wczytano pliki.";
        }

        private void LoadScheduleWindow() {
            Form f = new Form() { Text = "Wklej Grafik (Ctrl+V)", Size = new System.Drawing.Size(600, 400) };
            DataGridView dgv = new DataGridView() { Dock = DockStyle.Fill, AllowUserToAddRows = false };
            Button btnSave = new Button() { Text = "Zapisz", Dock = DockStyle.Bottom, Height = 40 };
            dgv.KeyDown += (s, e) => { if (e.Control && e.KeyCode == Keys.V) PasteToDgv(dgv); };
            btnSave.Click += (s, e) => { ProcessSchedule(dgv); f.Close(); };
            f.Controls.Add(dgv); f.Controls.Add(btnSave); f.ShowDialog();
        }

        private void PasteToDgv(DataGridView dgv) {
            string t = Clipboard.GetText(); if (string.IsNullOrEmpty(t)) return;
            dgv.Rows.Clear(); dgv.Columns.Clear();
            string[] lines = t.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            string[] headers = lines[0].Split('\t');
            foreach (var h in headers) dgv.Columns.Add(h, h);
            for (int i = 1; i < lines.Length; i++) dgv.Rows.Add(lines[i].Split('\t'));
        }

        private void ProcessSchedule(DataGridView dgv) {
            staffSchedule.Clear();
            for (int r = 0; r < dgv.Rows.Count; r++) {
                string loc = dgv.Rows[r].Cells[0].Value?.ToString().ToUpper().Trim() ?? "";
                for (int c = 1; c < dgv.Columns.Count; c++) {
                    if (int.TryParse(dgv.Columns[c].HeaderText, out int day)) {
                        string person = dgv.Rows[r].Cells[c].Value?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(person)) staffSchedule[new ScheduleKey { Loc = loc, Day = day }] = person;
                    }
                }
            }
        }

        private void ShowSettingsWindow() {
            Form f = new Form() { Text = "Ustawienia UPS", Size = new System.Drawing.Size(300, 200) };
            TextBox t1 = new TextBox() { Text = upsLicense, Dock = DockStyle.Top };
            TextBox t2 = new TextBox() { Text = upsUser, Dock = DockStyle.Top };
            TextBox t3 = new TextBox() { Text = upsPass, Dock = DockStyle.Top, UseSystemPasswordChar = true };
            Button b = new Button() { Text = "Zapisz", Dock = DockStyle.Bottom };
            b.Click += (s, e) => {
                upsLicense = t1.Text; upsUser = t2.Text; upsPass = t3.Text;
                File.WriteAllLines(settingsPath, new[] { upsLicense, upsUser, upsPass });
                f.Close();
            };
            f.Controls.Add(t3); f.Controls.Add(new Label { Text = "Hasło:", Dock = DockStyle.Top });
            f.Controls.Add(t2); f.Controls.Add(new Label { Text = "User ID:", Dock = DockStyle.Top });
            f.Controls.Add(t1); f.Controls.Add(new Label { Text = "Licencja:", Dock = DockStyle.Top });
            f.Controls.Add(b); f.ShowDialog();
        }

        private async void BtnRun_Click(object sender, EventArgs e) {
            btnRun.Enabled = false; apiSuccess = 0; apiFailed = 0;
            try { await GenerateReportAsync(); MessageBox.Show("Raport wygenerowany!"); }
            catch (Exception ex) { MessageBox.Show("Błąd: " + ex.Message); }
            finally { btnRun.Enabled = true; lblStatus.Text = "Gotowe."; }
        }

        private async System.Threading.Tasks.Task GenerateReportAsync()
        {
            var selectedLocs = clbWarehouses.CheckedItems.Cast<string>().ToList();
            var data = new Dictionary<string, SortedDictionary<DateTime, string>>();
            var dates = sortedFiles.Select(x => x.Date).ToList();
            DateTime lastDay = dates.Max();
            List<string> failedPackages = new List<string>();
            var personMissedScans = new Dictionary<string, int>();

            foreach (var f in sortedFiles) {
                using (var wb = new XLWorkbook(f.Path)) {
                    var ws = wb.Worksheets.FirstOrDefault();
                    if (ws == null) continue;
                    foreach (var row in ws.RangeUsed().RowsUsed().Skip(1)) {
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
                var wsStat = report.Worksheets.Add("Statystyki");
                ws.Cell(1, 1).Value = "Package ID";
                for (int i = 0; i < dates.Count; i++) ws.Cell(1, i + 2).Value = dates[i].ToShortDateString();
                
                int colStatus = dates.Count + 2;
                int colCity = dates.Count + 3;
                ws.Cell(1, colStatus).Value = "Status UPS";
                ws.Cell(1, colCity).Value = "Lokalizacja UPS";

                int r = 2;
                foreach (var pkg in data) {
                    var first = pkg.Value.Keys.Min();
                    var lastScan = pkg.Value.Keys.Max();
                    bool missingLast = !pkg.Value.ContainsKey(lastDay) && (lastDay - lastScan).TotalDays <= 3;

                    if (missingLast || pkg.Value.Count < dates.Count(d => d >= first && d <= lastScan)) {
                        ws.Cell(r, 1).Value = pkg.Key;
                        bool isActuallyOutside = false;

                        if (missingLast && chkEnableUPS.Checked && !string.IsNullOrEmpty(upsLicense)) {
                            lblStatus.Text = "UPS API: " + pkg.Key + "..."; Application.DoEvents();
                            var res = await GetUpsTracking(pkg.Key);
                            ws.Cell(r, colStatus).Value = res.Item1;
                            ws.Cell(r, colCity).Value = res.Item2;
                            string cityNorm = res.Item2.ToUpper();
                            if (!string.IsNullOrEmpty(res.Item2) && res.Item2 != "---" && !cityNorm.Contains("STRYKOW") && !cityNorm.Contains("DOBRA")) {
                                isActuallyOutside = true;
                            }
                        }

                        for (int i = 0; i < dates.Count; i++) {
                            DateTime d = dates[i];
                            if (pkg.Value.ContainsKey(d)) ws.Cell(r, i + 2).Value = pkg.Value[d];
                            else if (d > first) {
                                var cell = ws.Cell(r, i + 2);
                                if (isActuallyOutside && d == lastDay) {
                                    cell.Value = "DORĘCZONA/WYDANA"; cell.Style.Fill.BackgroundColor = XLColor.Green; cell.Style.Font.FontColor = XLColor.White;
                                } else {
                                    cell.Value = "BRAK SKANU"; cell.Style.Fill.BackgroundColor = XLColor.Salmon;
                                    var key = new ScheduleKey { Loc = pkg.Value[first], Day = d.Day };
                                    if (staffSchedule.TryGetValue(key, out string p)) {
                                        cell.CreateComment().AddText(p);
                                        if (!personMissedScans.ContainsKey(p)) personMissedScans[p] = 0;
                                        personMissedScans[p]++;
                                    }
                                }
                            }
                        }
                        r++;
                    }
                }
                wsStat.Cell(1, 1).Value = "Ranking braków:";
                int sr = 2;
                foreach (var kvp in personMissedScans.OrderByDescending(x => x.Value)) {
                    wsStat.Cell(sr, 1).Value = kvp.Key; wsStat.Cell(sr, 2).Value = kvp.Value; sr++;
                }
                ws.Columns().AdjustToContents();
                report.SaveAs(Path.Combine(selectedFolderPath, "Raport_AK0_" + DateTime.Now.ToString("ddMMyy_HHmm") + ".xlsx"));
            }
        }

        private async System.Threading.Tasks.Task<Tuple<string, string>> GetUpsTracking(string trackNum) {
            try {
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                string xml = "<?xml version=\"1.0\"?><AccessRequest><AccessLicenseNumber>" + upsLicense + "</AccessLicenseNumber><UserId>" + upsUser + "</UserId><Password>" + upsPass + "</Password></AccessRequest>" +
                             "<?xml version=\"1.0\"?><TrackRequest><Request><RequestAction>Track</RequestAction></Request><TrackingNumber>" + trackNum + "</TrackingNumber></TrackRequest>";
                using (var client = new HttpClient()) {
                    var resp = await client.PostAsync("https://www.ups.com/ups.app/xml/Track", new StringContent(xml, Encoding.UTF8, "application/x-www-form-urlencoded"));
                    var doc = XDocument.Parse(await resp.Content.ReadAsStringAsync());
                    var pkg = doc.Descendants("Package").FirstOrDefault();
                    if (pkg != null) {
                        var act = pkg.Descendants("Activity").FirstOrDefault();
                        string st = act?.Descendants("Status")?.FirstOrDefault()?.Descendants("Description")?.FirstOrDefault()?.Value ?? "Brak";
                        string ct = act?.Descendants("ActivityLocation")?.FirstOrDefault()?.Descendants("Address")?.FirstOrDefault()?.Descendants("City")?.FirstOrDefault()?.Value ?? "Nieznane";
                        apiSuccess++; return new Tuple<string, string>(st, ct);
                    }
                }
            } catch { }
            apiFailed++; return new Tuple<string, string>("Błąd API", "---");
        }

        [STAThread] static void Main() { Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false); Application.Run(new MainForm()); }
    }
}
