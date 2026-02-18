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
using ClosedXML.Excel; // Wymaga zainstalowania NuGet: ClosedXML

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

        // Klasy pomocnicze dla starszego .NET
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

        // ... [Reszta metod SelectFolder, LoadSchedule, LoadSettings pozostaje identyczna jak wcześniej] ...
        // [Poniżej kluczowa zmiana w logice GenerateReport dla .NET 4.8]

        private async System.Threading.Tasks.Task GenerateReportAsync()
        {
            var selectedLocs = clbWarehouses.CheckedItems.Cast<string>().ToList();
            var data = new Dictionary<string, SortedDictionary<DateTime, string>>();
            var dates = sortedFiles.Select(x => x.Date).ToList();
            DateTime lastDay = dates.Max();
            List<string> failedPackages = new List<string>();
            var personMissedScans = new Dictionary<string, int>();

            // Wczytywanie danych
            foreach (var f in sortedFiles) {
                using (var wb = new XLWorkbook(f.Path)) {
                    var ws = wb.Worksheets.FirstOrDefault(w => w.Name.ToUpper() == "AK0") ?? wb.Worksheets.FirstOrDefault();
                    if (ws == null) continue;
                    var used = ws.RangeUsed(); if (used == null) continue;
                    foreach (var row in used.RowsUsed().Skip(1)) {
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
                    bool hasGap = false;
                    for (int i = 0; i < dates.Count; i++) {
                        if (dates[i] > first && dates[i] < lastScan && !pkg.Value.ContainsKey(dates[i])) {
                            var next = pkg.Value.Keys.Where(d => d > dates[i]).Min();
                            if ((next - dates[i]).TotalDays <= 3) { hasGap = true; break; }
                        }
                    }
                    bool missingLast = !pkg.Value.ContainsKey(lastDay) && (lastDay - lastScan).TotalDays <= 3;

                    if (hasGap || missingLast) {
                        ws.Cell(r, 1).Value = pkg.Key;
                        bool isActuallyOutside = false;

                        if (missingLast && chkEnableUPS.Checked && !string.IsNullOrEmpty(upsLicense)) {
                            lblStatus.Text = "UPS API: " + pkg.Key + "..."; Application.DoEvents();
                            var res = await GetUpsTracking(pkg.Key);
                            ws.Cell(r, colStatus).Value = res.Item1;
                            ws.Cell(r, colCity).Value = res.Item2;

                            string cityNorm = res.Item2.ToUpper();
                            if (!string.IsNullOrEmpty(res.Item2) && res.Item2 != "---" && 
                                !cityNorm.Contains("STRYKOW") && !cityNorm.Contains("DOBRA")) {
                                isActuallyOutside = true;
                            }
                            if (res.Item1.Contains("Błąd")) failedPackages.Add(pkg.Key);
                        }

                        for (int i = 0; i < dates.Count; i++) {
                            DateTime d = dates[i];
                            if (pkg.Value.ContainsKey(d)) ws.Cell(r, i + 2).Value = pkg.Value[d];
                            else if (d > first) {
                                var cell = ws.Cell(r, i + 2);
                                if (isActuallyOutside && d == lastDay) {
                                    cell.Value = "DORĘCZONA/WYDANA"; 
                                    cell.Style.Fill.BackgroundColor = XLColor.Green;
                                    cell.Style.Font.FontColor = XLColor.White;
                                } else {
                                    cell.Value = "BRAK SKANU"; cell.Style.Fill.BackgroundColor = XLColor.Salmon;
                                    var key = new ScheduleKey { Loc = pkg.Value[first], Day = d.Day };
                                    if (staffSchedule.TryGetValue(key, out string person)) {
                                        cell.CreateComment().AddText(person);
                                        if (!personMissedScans.ContainsKey(person)) personMissedScans[person] = 0;
                                        personMissedScans[person]++;
                                    }
                                }
                            }
                        }
                        r++;
                    }
                }
                // ... [Kod statystyk i zapisu pliku identyczny] ...
                ws.Columns().AdjustToContents(); wsStat.Columns().AdjustToContents();
                report.SaveAs(Path.Combine(selectedFolderPath, "Raport_AK0_" + DateTime.Now.ToString("ddMMyy_HHmm") + ".xlsx"));
            }
        }

        private async System.Threading.Tasks.Task<Tuple<string, string>> GetUpsTracking(string trackNum)
        {
            try {
                // W .NET Framework 4.8 musimy wymusić TLS 1.2 dla nowoczesnych API
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                
                string xml = "<?xml version=\"1.0\"?><AccessRequest><AccessLicenseNumber>" + upsLicense + "</AccessLicenseNumber><UserId>" + upsUser + "</UserId><Password>" + upsPass + "</Password></AccessRequest>" +
                             "<?xml version=\"1.0\"?><TrackRequest><Request><RequestAction>Track</RequestAction></Request><TrackingNumber>" + trackNum + "</TrackingNumber></TrackRequest>";
                
                using (var client = new HttpClient()) {
                    var resp = await client.PostAsync("https://www.ups.com/ups.app/xml/Track", new StringContent(xml, Encoding.UTF8, "application/x-www-form-urlencoded"));
                    var content = await resp.Content.ReadAsStringAsync();
                    var doc = XDocument.Parse(content);
                    var package = doc.Descendants("Package").FirstOrDefault();
                    if (package != null) {
                        var activity = package.Descendants("Activity").FirstOrDefault();
                        string desc = activity?.Descendants("Status")?.FirstOrDefault()?.Descendants("StatusType")?.FirstOrDefault()?.Descendants("Description")?.FirstOrDefault()?.Value ?? "Brak opisu";
                        string city = activity?.Descendants("ActivityLocation")?.FirstOrDefault()?.Descendants("Address")?.FirstOrDefault()?.Descendants("City")?.FirstOrDefault()?.Value ?? "Nieznane";
                        apiSuccess++; return Tuple.Create(desc, city);
                    }
                }
            } catch { }
            apiFailed++; return Tuple.Create("Błąd API", "---");
        }

        [STAThread] static void Main() { Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false); Application.Run(new MainForm()); }
    }
}
