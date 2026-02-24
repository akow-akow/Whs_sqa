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
        private Button btnRun, btnSelectFolder, btnLoadSchedule, btnSettings, btnLoadReleased;
        private CheckBox chkEnableUPS, chkFilterI, chkFilterE;
        private Label lblStatus;
        private List<FileItem> sortedFiles;
        private HashSet<string> allDetectedLocs = new HashSet<string>();
        private string selectedFolderPath = "";
        private Dictionary<ScheduleKey, string> staffSchedule = new Dictionary<ScheduleKey, string>();
        private HashSet<string> releasedPackages = new HashSet<string>();

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
            this.Text = "AK0 Warehouse Scan Quality Analyzer";
            this.Size = new System.Drawing.Size(550, 920);
            this.StartPosition = FormStartPosition.CenterScreen;
            try { if (File.Exists("icon.ico")) this.Icon = new System.Drawing.Icon("icon.ico"); } catch { }

            FlowLayoutPanel topPanel = new FlowLayoutPanel() { Dock = DockStyle.Top, Height = 280, Padding = new Padding(10) };
            
            btnSelectFolder = new Button() { Text = "📁 1. WYBIERZ FOLDER AK0", Size = new System.Drawing.Size(245, 60), BackColor = System.Drawing.Color.LightSkyBlue, FlatStyle = FlatStyle.Flat, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
            btnSelectFolder.Click += (s, e) => SelectFolder();
            
            btnLoadSchedule = new Button() { Text = "📅 2a. WCZYTAJ GRAFIK", Size = new System.Drawing.Size(245, 60), BackColor = System.Drawing.Color.NavajoWhite, FlatStyle = FlatStyle.Flat, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
            btnLoadSchedule.Click += (s, e) => LoadScheduleWindow();

            btnLoadReleased = new Button() { Text = "🚚 2b. PRZESYŁKI ZWOLNIONE (DAT/TEKST)", Size = new System.Drawing.Size(500, 45), BackColor = System.Drawing.Color.LightSteelBlue, FlatStyle = FlatStyle.Flat, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
            btnLoadReleased.Click += (s, e) => LoadReleasedWindow();
            
            btnSettings = new Button() { Text = "⚙️ USTAWIENIA UPS API", Size = new System.Drawing.Size(500, 40), BackColor = System.Drawing.Color.LightGray, FlatStyle = FlatStyle.Flat };
            btnSettings.Click += (s, e) => ShowSettingsWindow();

            GroupBox gpFilters = new GroupBox() { Text = "Filtry magazynów (Początek nazwy)", Size = new System.Drawing.Size(500, 50) };
            chkFilterI = new CheckBox() { Text = "Import (I...)", Checked = true, AutoSize = true, Location = new System.Drawing.Point(10, 20) };
            chkFilterE = new CheckBox() { Text = "Export (E...)", Checked = true, AutoSize = true, Location = new System.Drawing.Point(150, 20) };
            chkFilterI.CheckedChanged += (s, e) => ApplyLocFilter();
            chkFilterE.CheckedChanged += (s, e) => ApplyLocFilter();
            gpFilters.Controls.Add(chkFilterI); gpFilters.Controls.Add(chkFilterE);

            topPanel.Controls.Add(btnSelectFolder);
            topPanel.Controls.Add(btnLoadSchedule);
            topPanel.Controls.Add(btnLoadReleased);
            topPanel.Controls.Add(btnSettings);
            topPanel.Controls.Add(gpFilters);

            clbWarehouses = new CheckedListBox() { Dock = DockStyle.Fill, CheckOnClick = true, Font = new System.Drawing.Font("Segoe UI", 10) };
            
            Panel pnlOptions = new Panel() { Dock = DockStyle.Bottom, Height = 40, BackColor = System.Drawing.Color.WhiteSmoke };
            chkEnableUPS = new CheckBox() { Text = "Automatyczna weryfikacja UPS API", AutoSize = true, Location = new System.Drawing.Point(10, 10), Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
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

        private void LoadReleasedWindow()
        {
            Form f = new Form() { Text = "Zarządzanie przesyłkami RELEASED", Size = new System.Drawing.Size(600, 500), StartPosition = FormStartPosition.CenterParent };
            Label lblInfo = new Label() { Text = "Wklej raport LUB wybierz plik WHOFILEXPT.DAT:", Dock = DockStyle.Top, Height = 30, TextAlign = System.Drawing.ContentAlignment.BottomLeft, Padding = new Padding(5) };
            TextBox txt = new TextBox() { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical, Font = new System.Drawing.Font("Consolas", 9) };
            Panel pnlButtons = new Panel() { Dock = DockStyle.Bottom, Height = 100 };
            Button btnFile = new Button() { Text = "📁 WYBIERZ PLIK WHOFILEXPT.DAT", Size = new System.Drawing.Size(570, 45), Location = new System.Drawing.Point(10, 5), BackColor = System.Drawing.Color.LightCyan, FlatStyle = FlatStyle.Flat };
            Button btnProcess = new Button() { Text = "✅ PRZETWÓRZ WKLEJONY TEKST", Size = new System.Drawing.Size(570, 40), Location = new System.Drawing.Point(10, 55), BackColor = System.Drawing.Color.LightSteelBlue, FlatStyle = FlatStyle.Flat };
            
            pnlButtons.Controls.Add(btnFile);
            pnlButtons.Controls.Add(btnProcess);

            Action<string[]> processLines = (lines) => {
                releasedPackages.Clear();
                foreach (var line in lines) {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    string[] parts = line.Split(',');
                    if (parts.Length > 5) {
                        string trackNum = parts[5].Trim();
                        if (!string.IsNullOrEmpty(trackNum)) releasedPackages.Add(trackNum);
                    }
                }
                MessageBox.Show($"Wczytano {releasedPackages.Count} unikalnych numerów paczek.");
                f.Close();
            };

            btnFile.Click += (s, e) => {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Pliki DAT (*.dat)|*.dat|Wszystkie pliki (*.*)|*.*" }) {
                    if (ofd.ShowDialog() == DialogResult.OK) processLines(File.ReadAllLines(ofd.FileName));
                }
            };
            btnProcess.Click += (s, e) => processLines(txt.Lines);

            f.Controls.Add(txt); f.Controls.Add(lblInfo); f.Controls.Add(pnlButtons);
            f.ShowDialog();
        }

        private void ApplyLocFilter() {
            clbWarehouses.Items.Clear();
            foreach (var loc in allDetectedLocs.OrderBy(x => x)) {
                bool isI = loc.StartsWith("I", StringComparison.OrdinalIgnoreCase);
                bool isE = loc.StartsWith("E", StringComparison.OrdinalIgnoreCase);
                if ((isI && chkFilterI.Checked) || (isE && chkFilterE.Checked)) clbWarehouses.Items.Add(loc);
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
            if (!Directory.Exists(selectedFolderPath)) return;
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
                try {
                    using (var wb = new XLWorkbook(f.Path)) {
                        var ws = wb.Worksheets.FirstOrDefault(w => w.Name.ToUpper().Contains("AK0")) ?? wb.Worksheets.FirstOrDefault();
                        var range = ws?.RangeUsed(); if (range == null) continue;
                        foreach (var row in range.RowsUsed().Skip(1)) {
                            string val = row.Cell(1).GetString().Trim();
                            if (!string.IsNullOrEmpty(val)) allDetectedLocs.Add(val);
                        }
                    }
                } catch { }
            }
            ApplyLocFilter();
            btnRun.Enabled = true;
            lblStatus.Text = "Wczytano " + sortedFiles.Count + " plików.";
        }

        private void LoadScheduleWindow() {
            Form f = new Form() { Text = "Wklej Grafik (Ctrl+V)", Size = new System.Drawing.Size(800, 500), StartPosition = FormStartPosition.CenterParent };
            DataGridView dgv = new DataGridView() { Dock = DockStyle.Fill, AllowUserToAddRows = false };
            Button btnSave = new Button() { Text = "Zapisz i Mapuj Grafik", Dock = DockStyle.Bottom, Height = 45, BackColor = System.Drawing.Color.PaleGreen };
            dgv.KeyDown += (s, e) => { if (e.Control && e.KeyCode == Keys.V) PasteToDgv(dgv); };
            btnSave.Click += (s, e) => { ProcessSchedule(dgv); f.Close(); };
            f.Controls.Add(dgv); f.Controls.Add(btnSave); f.ShowDialog();
        }

        private void PasteToDgv(DataGridView dgv) {
            string t = Clipboard.GetText(); if (string.IsNullOrEmpty(t)) return;
            dgv.Rows.Clear(); dgv.Columns.Clear();
            string[] lines = t.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            if (lines.Length == 0) return;
            string[] headers = lines[0].Split('\t');
            foreach (var h in headers) dgv.Columns.Add(h, h);
            for (int i = 1; i < lines.Length; i++) dgv.Rows.Add(lines[i].Split('\t'));
        }

        private void ProcessSchedule(DataGridView dgv) {
            staffSchedule.Clear();
            int mappedCount = 0;
            for (int r = 0; r < dgv.Rows.Count; r++) {
                string cellA = dgv.Rows[r].Cells[0].Value?.ToString().Trim().ToLower() ?? "";
                if (string.IsNullOrEmpty(cellA)) continue;

                List<string> mappedLocs = new List<string>();
                bool isSmalls = cellA.Contains("smalls");

                if (cellA.Contains("mag")) {
                    string num = Regex.Match(cellA, @"\d+").Value;
                    if (!string.IsNullOrEmpty(num)) {
                        mappedLocs.Add("IWMAGAZYN" + num);
                        mappedLocs.Add("EWMAGEXP" + num);
                    }
                } else if (isSmalls) {
                    mappedLocs.Add("IWMSMALLS"); mappedLocs.Add("IWMSMALLS1"); mappedLocs.Add("IWMSMALLSXX");
                }

                if (mappedLocs.Count > 0) {
                    for (int c = 1; c < dgv.Columns.Count; c++) {
                        string dayHeader = Regex.Match(dgv.Columns[c].HeaderText, @"\d+").Value;
                        if (int.TryParse(dayHeader, out int day)) {
                            string p1 = dgv.Rows[r].Cells[c].Value?.ToString().Trim() ?? "";
                            if (isSmalls && r + 1 < dgv.Rows.Count) {
                                string p2 = dgv.Rows[r + 1].Cells[c].Value?.ToString().Trim() ?? "";
                                if (!string.IsNullOrEmpty(p2)) p1 = string.IsNullOrEmpty(p1) ? p2 : p1 + " / " + p2;
                            }
                            if (!string.IsNullOrEmpty(p1)) {
                                foreach (var ml in mappedLocs) staffSchedule[new ScheduleKey { Loc = ml.ToUpper(), Day = day }] = p1;
                            }
                        }
                    }
                    mappedCount++;
                    if (isSmalls) r++; 
                }
            }
            MessageBox.Show($"Zmapowano grafik dla {mappedCount} pozycji.");
        }

        private async void BtnRun_Click(object sender, EventArgs e) {
            btnRun.Enabled = false;
            try { await GenerateReportAsync(); MessageBox.Show("Raport wygenerowany!"); }
            catch (Exception ex) { MessageBox.Show("Błąd: " + ex.Message); }
            finally { btnRun.Enabled = true; lblStatus.Text = "Gotowe."; }
        }

        private async System.Threading.Tasks.Task GenerateReportAsync() {
            var selectedLocs = clbWarehouses.CheckedItems.Cast<string>().ToList();
            var data = new Dictionary<string, SortedDictionary<DateTime, string>>();
            var pkgStartedInSelected = new HashSet<string>();
            
            var dates = sortedFiles.Select(x => x.Date).ToList();
            DateTime lastDay = dates.Max();

            // KROK 1: Identyfikacja paczek, które były w zaznaczonych magazynach
            foreach (var f in sortedFiles) {
                using (var wb = new XLWorkbook(f.Path)) {
                    var ws = wb.Worksheets.FirstOrDefault(w => w.Name.ToUpper().Contains("AK0")) ?? wb.Worksheets.FirstOrDefault();
                    var range = ws?.RangeUsed(); if (range == null) continue;
                    foreach (var row in range.RowsUsed().Skip(1)) {
                        string l = row.Cell(1).GetString().Trim();
                        string p = row.Cell(2).GetString().Trim();
                        if (selectedLocs.Contains(l)) pkgStartedInSelected.Add(p);
                    }
                }
            }

            // KROK 2: Budowanie historii dla tych paczek
            foreach (var f in sortedFiles) {
                using (var wb = new XLWorkbook(f.Path)) {
                    var ws = wb.Worksheets.FirstOrDefault(w => w.Name.ToUpper().Contains("AK0")) ?? wb.Worksheets.FirstOrDefault();
                    var range = ws?.RangeUsed(); if (range == null) continue;
                    foreach (var row in range.RowsUsed().Skip(1)) {
                        string l = row.Cell(1).GetString().Trim();
                        string p = row.Cell(2).GetString().Trim();
                        if (pkgStartedInSelected.Contains(p)) {
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
                
                int colStatus = dates.Count + 2;
                int colCity = dates.Count + 3;
                int colStaff = dates.Count + 4;
                ws.Cell(1, colStatus).Value = "Status UPS";
                ws.Cell(1, colCity).Value = "Lokalizacja UPS";
                ws.Cell(1, colStaff).Value = "Osoba Odpowiedzialna";

                int r = 2;
                foreach (var pkg in data) {
                    DateTime first = pkg.Value.Keys.Min();
                    bool isMissingTodayInSelected = !pkg.Value.ContainsKey(lastDay) || !selectedLocs.Contains(pkg.Value[lastDay]);
                    
                    bool hasGaps = false;
                    for(DateTime d = first; d <= lastDay; d = d.AddDays(1)) {
                        var targetDate = dates.FirstOrDefault(dt => dt.Date == d.Date);
                        if (targetDate != default(DateTime) && !pkg.Value.ContainsKey(targetDate)) { hasGaps = true; break; }
                    }

                    if (isMissingTodayInSelected || hasGaps) {
                        ws.Cell(r, 1).Value = pkg.Key;
                        bool isActuallyOutside = false;
                        bool isReleased = releasedPackages.Contains(pkg.Key);
                        bool isOutForDelivery = false; // Nowy warunek z API
                        bool existsAnywhereToday = pkg.Value.ContainsKey(lastDay);

                        if (!existsAnywhereToday && chkEnableUPS.Checked && !string.IsNullOrEmpty(upsLicense)) {
                            lblStatus.Text = "UPS: " + pkg.Key + "..."; Application.DoEvents();
                            var res = await GetUpsTracking(pkg.Key);
                            ws.Cell(r, colStatus).Value = res.Item1;
                            ws.Cell(r, colCity).Value = res.Item2;

                            // LOGIKA: Jeśli "Out For Delivery", traktujemy jak Released (nawet w Strykowie/Dobrej)
                            if (res.Item1.ToUpper().Contains("OUT FOR DELIVERY")) {
                                isOutForDelivery = true;
                            }
                            else if (!string.IsNullOrEmpty(res.Item2) && !res.Item2.ToUpper().Contains("STRYKOW") && !res.Item2.ToUpper().Contains("DOBRA")) {
                                isActuallyOutside = true;
                            }
                        }

                        for (int i = 0; i < dates.Count; i++) {
                            DateTime d = dates[i];
                            if (pkg.Value.ContainsKey(d)) {
                                string loc = pkg.Value[d];
                                ws.Cell(r, i + 2).Value = loc;
                                if (!selectedLocs.Contains(loc)) ws.Cell(r, i + 2).Style.Fill.BackgroundColor = XLColor.LightGray;
                            }
                            else if (d > first) {
                                var cell = ws.Cell(r, i + 2);
                                
                                // Oznaczanie jako Released / Out For Delivery / Doręczona
                                if (d == lastDay && (isReleased || isOutForDelivery)) {
                                    cell.Value = isOutForDelivery ? "OUT FOR DELIVERY" : "RELEASED";
                                    cell.Style.Fill.BackgroundColor = XLColor.LightSkyBlue;
                                } 
                                else if (isActuallyOutside && d == lastDay) {
                                    cell.Value = "DORĘCZONA"; cell.Style.Fill.BackgroundColor = XLColor.Green; cell.Style.Font.FontColor = XLColor.White;
                                } 
                                else {
                                    cell.Value = "BRAK SKANU"; cell.Style.Fill.BackgroundColor = XLColor.Salmon;
                                    string lastKnownLoc = pkg.Value.Where(kv => kv.Key < d).OrderByDescending(kv => kv.Key).FirstOrDefault().Value ?? "";
                                    var key = new ScheduleKey { Loc = lastKnownLoc.ToUpper(), Day = d.Day };
                                    if (staffSchedule.TryGetValue(key, out string pStaff)) {
                                        cell.CreateComment().AddText(pStaff);
                                        ws.Cell(r, colStaff).Value = pStaff;
                                    }
                                }
                            }
                        }
                        r++;
                    }
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
                        string st = act?.Descendants("Status")?.FirstOrDefault()?.Descendants("StatusType")?.FirstOrDefault()?.Descendants("Description")?.FirstOrDefault()?.Value ?? "Brak";
                        string ct = act?.Descendants("ActivityLocation")?.FirstOrDefault()?.Descendants("Address")?.FirstOrDefault()?.Descendants("City")?.FirstOrDefault()?.Value ?? "Nieznane";
                        return new Tuple<string, string>(st, ct);
                    }
                }
            } catch { }
            return new Tuple<string, string>("Błąd API", "---");
        }

        private void LoadSettings() { if (File.Exists(settingsPath)) { var lines = File.ReadAllLines(settingsPath); if (lines.Length >= 3) { upsLicense = lines[0]; upsUser = lines[1]; upsPass = lines[2]; } } }
        private void ShowSettingsWindow() {
            Form f = new Form() { Text = "Ustawienia UPS", Size = new System.Drawing.Size(300, 250), StartPosition = FormStartPosition.CenterParent };
            TextBox t1 = new TextBox() { Text = upsLicense, Dock = DockStyle.Top };
            TextBox t2 = new TextBox() { Text = upsUser, Dock = DockStyle.Top };
            TextBox t3 = new TextBox() { Text = upsPass, Dock = DockStyle.Top, UseSystemPasswordChar = true };
            Button b = new Button() { Text = "Zapisz", Dock = DockStyle.Bottom, Height = 40 };
            b.Click += (s, e) => { upsLicense = t1.Text; upsUser = t2.Text; upsPass = t3.Text; File.WriteAllLines(settingsPath, new[] { upsLicense, upsUser, upsPass }); f.Close(); };
            f.Controls.Add(t3); f.Controls.Add(new Label { Text = "Hasło UPS:", Dock = DockStyle.Top, Height = 25 });
            f.Controls.Add(t2); f.Controls.Add(new Label { Text = "User ID:", Dock = DockStyle.Top, Height = 25 });
            f.Controls.Add(t1); f.Controls.Add(new Label { Text = "Access License Number:", Dock = DockStyle.Top, Height = 25 });
            f.Controls.Add(b); f.ShowDialog();
        }

        [STAThread] static void Main() { Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false); Application.Run(new MainForm()); }
    }
}
