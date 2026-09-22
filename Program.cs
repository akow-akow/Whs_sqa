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
        // Kontrolka zakładek dzieląca aplikację na dwa tryby
        private TabControl tabControlMain;
        private TabPage tabAk0Analyzer;
        private TabPage tabUnloadAnalyzer;

        // Kontrolki - Zakładka 1: AK0 Analyzer (Oryginał)
        private CheckedListBox clbWarehouses;
        private Button btnRun, btnSelectFolder, btnLoadSchedule, btnSettings, btnLoadReleased, btnLoadPostcodes;
        private CheckBox chkEnableUPS, chkFilterI, chkFilterE;
        private Label lblStatus;
        private List<FileItem> sortedFiles;
        private HashSet<string> allDetectedLocs = new HashSet<string>();
        private string selectedFolderPath = "";
        private Dictionary<ScheduleKey, string> staffSchedule = new Dictionary<ScheduleKey, string>();
        private HashSet<string> releasedPackages = new HashSet<string>();
        private Dictionary<string, string> postcodeMap = new Dictionary<string, string>();

        // Kontrolki - Zakładka 2: Rozładunki / Boxy UPS (Nowość)
        private Button btnSelectUnloadFolder, btnSelectAk0UnloadFolder, btnRunUnload;
        private Label lblUnloadStatus, lblUnloadFolderPath, lblAk0UnloadFolderPath;
        private string unloadFolderPath = "";
        private string ak0UnloadFolderPath = "";

        // Ustawienia UPS (w tym konfigurowalny URL API)
        private string upsLicense = "", upsUser = "", upsPass = "", upsApiUrl = "https://www.ups.com/ups.app/xml/Track";
        private readonly string settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ups_settings.ini");
        private readonly string defaultPostcodePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "postcode.xml");

        struct FileItem { public string Path; public DateTime Date; }
        struct ScheduleKey { 
            public string Loc; public int Day;
            public override bool Equals(object obj) => obj is ScheduleKey other && Loc == other.Loc && Day == other.Day;
            public override int GetHashCode() => (Loc?.GetHashCode() ?? 0) ^ Day.GetHashCode();
        }

        public MainForm()
        {
            LoadSettings();
            this.Text = "AK0 Warehouse & Unload Quality Analyzer";
            this.Size = new System.Drawing.Size(600, 950);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Główny TabControl
            tabControlMain = new TabControl() { Dock = DockStyle.Fill };
            
            tabAk0Analyzer = new TabPage("Analiza AK0 (Oryginał)");
            tabUnloadAnalyzer = new TabPage("Rozładunki / Boxy UPS (Nowość)");

            BuildAk0Tab();
            BuildUnloadTab();

            tabControlMain.TabPages.Add(tabAk0Analyzer);
            tabControlMain.TabPages.Add(tabUnloadAnalyzer);

            this.Controls.Add(tabControlMain);

            // Automatyczne wczytanie kodów przy starcie
            AutoLoadPostcode();
        }

        private void BuildAk0Tab()
        {
            FlowLayoutPanel topPanel = new FlowLayoutPanel() { Dock = DockStyle.Top, Height = 330, Padding = new Padding(10) };
            
            btnSelectFolder = new Button() { Text = "📁 1. WYBIERZ FOLDER AK0", Size = new System.Drawing.Size(265, 60), BackColor = System.Drawing.Color.LightSkyBlue, FlatStyle = FlatStyle.Flat, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
            btnSelectFolder.Click += (s, e) => SelectFolder();
            
            btnLoadSchedule = new Button() { Text = "📅 2a. WCZYTAJ GRAFIK", Size = new System.Drawing.Size(265, 60), BackColor = System.Drawing.Color.NavajoWhite, FlatStyle = FlatStyle.Flat, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
            btnLoadSchedule.Click += (s, e) => LoadScheduleWindow();

            btnLoadReleased = new Button() { Text = "🚚 2b. PRZESYŁKI ZWOLNIONE (WIELE PLIKÓW DAT)", Size = new System.Drawing.Size(540, 45), BackColor = System.Drawing.Color.LightSteelBlue, FlatStyle = FlatStyle.Flat, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
            btnLoadReleased.Click += (s, e) => LoadReleasedWindow();

            btnLoadPostcodes = new Button() { Text = "🗺️ 2c. WCZYTAJ POSTCODE.XML (RĘCZNIE)", Size = new System.Drawing.Size(540, 45), BackColor = System.Drawing.Color.Thistle, FlatStyle = FlatStyle.Flat, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
            btnLoadPostcodes.Click += (s, e) => LoadPostcodeXml(null);
            
            btnSettings = new Button() { Text = "⚙️ USTAWIENIA UPS API & URL", Size = new System.Drawing.Size(540, 40), BackColor = System.Drawing.Color.LightGray, FlatStyle = FlatStyle.Flat };
            btnSettings.Click += (s, e) => ShowSettingsWindow();

            GroupBox gpFilters = new GroupBox() { Text = "Filtry magazynów (Początek nazwy)", Size = new System.Drawing.Size(540, 50) };
            chkFilterI = new CheckBox() { Text = "Import (I...)", Checked = true, AutoSize = true, Location = new System.Drawing.Point(10, 20) };
            chkFilterE = new CheckBox() { Text = "Export (E...)", Checked = true, AutoSize = true, Location = new System.Drawing.Point(170, 20) };
            chkFilterI.CheckedChanged += (s, e) => ApplyLocFilter();
            chkFilterE.CheckedChanged += (s, e) => ApplyLocFilter();
            gpFilters.Controls.Add(chkFilterI); gpFilters.Controls.Add(chkFilterE);

            topPanel.Controls.Add(btnSelectFolder);
            topPanel.Controls.Add(btnLoadSchedule);
            topPanel.Controls.Add(btnLoadReleased);
            topPanel.Controls.Add(btnLoadPostcodes);
            topPanel.Controls.Add(btnSettings);
            topPanel.Controls.Add(gpFilters);

            clbWarehouses = new CheckedListBox() { Dock = DockStyle.Fill, CheckOnClick = true, Font = new System.Drawing.Font("Segoe UI", 10) };
            
            Panel pnlOptions = new Panel() { Dock = DockStyle.Bottom, Height = 40, BackColor = System.Drawing.Color.WhiteSmoke };
            chkEnableUPS = new CheckBox() { Text = "Automatyczna weryfikacja UPS API (Status + Kod)", AutoSize = true, Location = new System.Drawing.Point(10, 10), Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
            pnlOptions.Controls.Add(chkEnableUPS);

            btnRun = new Button() { Text = "🚀 3. GENERUJ RAPORT", Dock = DockStyle.Bottom, Height = 70, BackColor = System.Drawing.Color.LightGreen, Enabled = false, Font = new System.Drawing.Font("Segoe UI", 11, System.Drawing.FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btnRun.Click += BtnRun_Click;

            lblStatus = new Label() { Text = "Gotowy", Dock = DockStyle.Bottom, Height = 40, TextAlign = System.Drawing.ContentAlignment.MiddleCenter, BackColor = System.Drawing.Color.WhiteSmoke, BorderStyle = BorderStyle.FixedSingle };

            tabAk0Analyzer.Controls.Add(clbWarehouses);
            tabAk0Analyzer.Controls.Add(new Label() { Text = " Magazyny do analizy:", Dock = DockStyle.Top, Height = 25, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) });
            tabAk0Analyzer.Controls.Add(topPanel);
            tabAk0Analyzer.Controls.Add(pnlOptions);
            tabAk0Analyzer.Controls.Add(lblStatus);
            tabAk0Analyzer.Controls.Add(btnRun);
        }

        private void BuildUnloadTab()
        {
            Panel pnlUnloadTop = new Panel() { Dock = DockStyle.Top, Height = 320, Padding = new Padding(15) };

            btnSelectUnloadFolder = new Button() { Text = "📁 1. WYBIERZ FOLDER Z PLIKAMI ROZŁADUNKOWYMI", Size = new System.Drawing.Size(530, 50), Location = new System.Drawing.Point(15, 15), BackColor = System.Drawing.Color.Moccasin, FlatStyle = FlatStyle.Flat, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
            btnSelectUnloadFolder.Click += (s, e) => SelectUnloadFolder();

            lblUnloadFolderPath = new Label() { Text = "Brak wybranego folderu z plikami rozładunkowymi.", Size = new System.Drawing.Size(530, 40), Location = new System.Drawing.Point(15, 75), TextAlign = System.Drawing.ContentAlignment.MiddleLeft };

            btnSelectAk0UnloadFolder = new Button() { Text = "📁 2. WYBIERZ FOLDER Z PLIKAMI AK0 (DO WERYFIKACJI)", Size = new System.Drawing.Size(530, 50), Location = new System.Drawing.Point(15, 125), BackColor = System.Drawing.Color.LightSkyBlue, FlatStyle = FlatStyle.Flat, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) };
            btnSelectAk0UnloadFolder.Click += (s, e) => SelectAk0UnloadFolder();

            lblAk0UnloadFolderPath = new Label() { Text = "Brak wybranego folderu z plikami AK0.", Size = new System.Drawing.Size(530, 40), Location = new System.Drawing.Point(15, 185), TextAlign = System.Drawing.ContentAlignment.MiddleLeft };

            btnRunUnload = new Button() { Text = "🚀 3. GENERUJ RAPORT ROZŁADUNKÓW / BOXÓW", Size = new System.Drawing.Size(530, 60), Location = new System.Drawing.Point(15, 235), BackColor = System.Drawing.Color.LightGreen, Enabled = false, FlatStyle = FlatStyle.Flat, Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold) };
            btnRunUnload.Click += BtnRunUnload_Click;

            pnlUnloadTop.Controls.Add(btnSelectUnloadFolder);
            pnlUnloadTop.Controls.Add(lblUnloadFolderPath);
            pnlUnloadTop.Controls.Add(btnSelectAk0UnloadFolder);
            pnlUnloadTop.Controls.Add(lblAk0UnloadFolderPath);
            pnlUnloadTop.Controls.Add(btnRunUnload);

            lblUnloadStatus = new Label() { Text = "Gotowy do analizy rozładunków.", Dock = DockStyle.Bottom, Height = 45, TextAlign = System.Drawing.ContentAlignment.MiddleCenter, BackColor = System.Drawing.Color.WhiteSmoke, BorderStyle = BorderStyle.FixedSingle };

            tabUnloadAnalyzer.Controls.Add(pnlUnloadTop);
            tabUnloadAnalyzer.Controls.Add(lblUnloadStatus);
        }

        private void SelectUnloadFolder()
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog()) {
                if (fbd.ShowDialog() == DialogResult.OK) {
                    unloadFolderPath = fbd.SelectedPath;
                    lblUnloadFolderPath.Text = "Folder rozładunków: " + unloadFolderPath;
                    CheckUnloadReady();
                }
            }
        }

        private void SelectAk0UnloadFolder()
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog()) {
                if (fbd.ShowDialog() == DialogResult.OK) {
                    ak0UnloadFolderPath = fbd.SelectedPath;
                    lblAk0UnloadFolderPath.Text = "Folder AK0: " + ak0UnloadFolderPath;
                    CheckUnloadReady();
                }
            }
        }

        private void CheckUnloadReady()
        {
            if (!string.IsNullOrEmpty(unloadFolderPath) && !string.IsNullOrEmpty(ak0UnloadFolderPath)) {
                btnRunUnload.Enabled = true;
            }
        }

        private async void BtnRunUnload_Click(object sender, EventArgs e)
        {
            btnRunUnload.Enabled = false;
            try {
                lblUnloadStatus.Text = "Trwa przetwarzanie rozładunków i weryfikacja UPS API...";
                Application.DoEvents();
                await GenerateUnloadReportAsync();
                MessageBox.Show("Raport rozładunków został wygenerowany pomyślnie!");
            } catch (Exception ex) {
                MessageBox.Show("Błąd podczas generowania raportu rozładunków: " + ex.Message);
            } finally {
                btnRunUnload.Enabled = true;
                lblUnloadStatus.Text = "Gotowy.";
            }
        }

        private async System.Threading.Tasks.Task GenerateUnloadReportAsync()
        {
            // 1. Wczytanie plików AK0 z wybranego folderu dla rozładunków
            var ak0Files = Directory.GetFiles(ak0UnloadFolderPath, "*.xlsx");
            var ak0FileItems = new List<FileItem>();
            foreach (var f in ak0Files) {
                string fn = Path.GetFileName(f);
                var m = Regex.Match(fn, @"(\d{2}\.\d{2}\.\d{4})");
                if (m.Success && fn.ToUpper().StartsWith("AK0")) {
                    if (DateTime.TryParseExact(m.Value, "dd.MM.yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime dt))
                        ak0FileItems.Add(new FileItem { Path = f, Date = dt });
                }
            }
            ak0FileItems = ak0FileItems.OrderBy(x => x.Date).ToList();
            if (ak0FileItems.Count == 0) {
                throw new Exception("Nie znaleziono prawidłowych plików AK0 w wybranym folderze!");
            }

            // Mapowanie: Paczka -> Słownik (Data -> Lokalizacja w AK0)
            Dictionary<string, SortedDictionary<DateTime, string>> packageAk0History = new Dictionary<string, SortedDictionary<DateTime, string>>();
            foreach (var f in ak0FileItems) {
                using (var wb = new XLWorkbook(f.Path)) {
                    var ws = wb.Worksheets.FirstOrDefault(w => w.Name.ToUpper().Contains("AK0")) ?? wb.Worksheets.FirstOrDefault();
                    var range = ws?.RangeUsed(); if (range == null) continue;
                    foreach (var row in range.RowsUsed().Skip(1)) {
                        string l = row.Cell(1).GetString().Trim();
                        string p = row.Cell(2).GetString().Trim();
                        if (!string.IsNullOrEmpty(p)) {
                            if (!packageAk0History.ContainsKey(p)) packageAk0History[p] = new SortedDictionary<DateTime, string>();
                            packageAk0History[p][f.Date] = l;
                        }
                    }
                }
            }

            // 2. Wczytanie plików rozładunkowych z hubu
            var unloadFiles = Directory.GetFiles(unloadFolderPath, "*.xlsx");
            Dictionary<string, List<string>> boxPackagesMap = new Dictionary<string, List<string>>();

            foreach (var file in unloadFiles) {
                string fileName = Path.GetFileNameWithoutExtension(file);
                var matches = Regex.Matches(fileName, @"(UPST\w+)", RegexOptions.IgnoreCase);
                List<string> boxesInFile = new List<string>();
                foreach (Match m in matches) {
                    boxesInFile.Add(m.Value.ToUpper());
                }

                if (boxesInFile.Count == 0) continue;

                using (var wb = new XLWorkbook(file)) {
                    foreach (var ws in wb.Worksheets) {
                        string wsNameTrim = ws.Name.Trim().ToUpper();
                        string matchedBox = boxesInFile.FirstOrDefault(b => wsNameTrim.Contains(b)) ?? boxesInFile.First();

                        if (!boxPackagesMap.ContainsKey(matchedBox)) boxPackagesMap[matchedBox] = new List<string>();

                        var range = ws.RangeUsed();
                        if (range == null) continue;

                        foreach (var row in range.RowsUsed().Skip(1)) {
                            string pkg = row.Cell(1).GetString().Trim();
                            if (string.IsNullOrEmpty(pkg) || pkg.Length < 5) {
                                pkg = row.Cell(2).GetString().Trim();
                            }
                            if (!string.IsNullOrEmpty(pkg) && !boxPackagesMap[matchedBox].Contains(pkg)) {
                                boxPackagesMap[matchedBox].Add(pkg);
                            }
                        }
                    }
                }
            }

            if (boxPackagesMap.Count == 0) {
                throw new Exception("Nie znaleziono żadnych boxów ani paczek w plikach rozładunkowych!");
            }

            DateTime today = DateTime.Now.Date;

            // 3. Generowanie pliku Excel z wynikami
            using (var report = new XLWorkbook()) {
                var wsReport = report.Worksheets.Add("Rozładunki i Boxy");
                int colIndex = 1;

                foreach (var boxEntry in boxPackagesMap) {
                    string boxName = boxEntry.Key;
                    var pkgs = boxEntry.Value;

                    wsReport.Cell(1, colIndex).Value = $"Box: {boxName} ({pkgs.Count})";
                    wsReport.Cell(1, colIndex).Style.Font.Bold = true;
                    wsReport.Cell(1, colIndex).Style.Fill.BackgroundColor = XLColor.LightGray;

                    int rowIndex = 2;
                    foreach (string pkg in pkgs) {
                        var cell = wsReport.Cell(rowIndex, colIndex);
                        cell.Value = pkg;

                        // Sprawdzenie czy paczka jest w zwolnionych (z pliku .DAT)
                        bool isReleased = releasedPackages.Contains(pkg);

                        // Analiza obecności paczki w AK0
                        bool foundInAk0 = packageAk0History.ContainsKey(pkg);
                        DateTime lastSeenDate = default(DateTime);
                        string lastLoc = "";

                        if (foundInAk0) {
                            var history = packageAk0History[pkg];
                            lastSeenDate = history.Keys.Max();
                            lastLoc = history[lastSeenDate];
                        }

                        // Warunek powrotu do GB (EWMAGCFRTS)
                        bool returnedToGb = false;
                        if (foundInAk0) {
                            foreach (var kv in packageAk0History[pkg]) {
                                if (kv.Value.Equals("EWMAGCFRTS", StringComparison.OrdinalIgnoreCase)) {
                                    returnedToGb = true; 
                                }
                            }
                            if (lastLoc.Equals("EWMAGCFRTS", StringComparison.OrdinalIgnoreCase)) returnedToGb = true;
                        }

                        bool recentInAk0 = foundInAk0 && (today - lastSeenDate).TotalDays <= 3;

                        // Dodatkowa weryfikacja przez UPS API (dla paczek nieobecnych w AK0 od ponad 3 dni lub nierozpoznanych)
                        bool isOkByUps = false;
                        string upsStatusInfo = "";
                        string upsDateLoc = "";
                        string apiLastCity = "";
                        string apiDescription = "";

                        if (!recentInAk0 && !returnedToGb && !isReleased) {
                            if (!string.IsNullOrEmpty(upsLicense)) {
                                lblUnloadStatus.Text = $"Weryfikacja UPS API dla paczki: {pkg}...";
                                Application.DoEvents();
                                var upsRes = await GetUpsTracking(pkg);
                                apiDescription = upsRes.Item1;
                                apiLastCity = upsRes.Item2;
                                string statusDesc = apiDescription.ToUpper();
                                string city = apiLastCity;

                                bool isDelivered = statusDesc.Contains("DELIVERED") || statusDesc.Contains("DORĘCZONA");
                                bool isOutForDelivery = statusDesc.Contains("OUT FOR DELIVERY");
                                bool isOutsideStrykowPoland = !string.IsNullOrEmpty(city) && 
                                                              !city.ToUpper().Contains("STRYKOW") && 
                                                              !city.ToUpper().Contains("DOBRA") && 
                                                              !city.ToUpper().Contains("NIEZNANE") &&
                                                              !statusDesc.Contains("BŁĄD");

                                if (isDelivered || isOutForDelivery || isOutsideStrykowPoland) {
                                    isOkByUps = true;
                                    upsStatusInfo = isDelivered ? "Doręczone" : (isOutForDelivery ? "OUT FOR DELIVERY" : $"W drodze ({city})");
                                    upsDateLoc = $"{upsStatusInfo} - {DateTime.Now:dd-MM-yyyy}";
                                }
                            }
                        }

                        // Ocena końcowa statusu wiersza w raporcie boxów
                        if (recentInAk0 || returnedToGb || isReleased || isOkByUps) {
                            // Status OK -> zielony / błękitny lub brak czerwonego
                            if (isReleased) {
                                cell.Style.Fill.BackgroundColor = XLColor.LightSkyBlue;
                                cell.CreateComment().AddText("Przesyłka zwolniona (z plików .DAT)");
                            } else if (isOkByUps) {
                                cell.Style.Fill.BackgroundColor = XLColor.LightGreen;
                                cell.CreateComment().AddText($"UPS OK: {upsDateLoc}");
                            } else if (returnedToGb) {
                                cell.CreateComment().AddText($"Zwrot do GB (EWMAGCFRTS) - Ostatnio: {lastSeenDate:dd-MM-yyyy}");
                            } else {
                                cell.CreateComment().AddText($"Obecna w AK0: {lastLoc} ({lastSeenDate:dd-MM-yyyy})");
                            }
                        } else {
                            // Problem / Brak w AK0 > 3 dni i brak potwierdzenia UPS/zwolnienia
                            cell.Style.Fill.BackgroundColor = XLColor.Salmon;
                            string commentText = "";
                            if (!foundInAk0) {
                                if (!string.IsNullOrEmpty(apiLastCity) && apiLastCity != "---") {
                                    commentText = $"Brak w AK0 | API: {apiLastCity} - {apiDescription}";
                                } else {
                                    commentText = "Brak w AK0 / Nieznana";
                                }
                            } else {
                                commentText = $"{lastLoc} {lastSeenDate:dd-MM-yyyy}";
                            }
                            cell.CreateComment().AddText(commentText);
                        }

                        rowIndex++;
                    }
                    colIndex++;
                }

                wsReport.Columns().AdjustToContents();
                string outPath = Path.Combine(unloadFolderPath, "Raport_Rozladunki_Boxy_" + DateTime.Now.ToString("ddMMyy_HHmm") + ".xlsx");
                report.SaveAs(outPath);
            }
        }

        private void AutoLoadPostcode() {
            if (File.Exists(defaultPostcodePath)) {
                LoadPostcodeXml(defaultPostcodePath);
            }
        }

        private void LoadPostcodeXml(string path)
        {
            string fileToLoad = path;
            if (string.IsNullOrEmpty(fileToLoad)) {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Pliki XML (*.xml)|*.xml" }) {
                    if (ofd.ShowDialog() == DialogResult.OK) fileToLoad = ofd.FileName;
                }
            }

            if (!string.IsNullOrEmpty(fileToLoad) && File.Exists(fileToLoad)) {
                try {
                    var doc = XDocument.Load(fileToLoad);
                    postcodeMap.Clear();
                    foreach (var pc in doc.Descendants("postcode")) {
                        string code = pc.Element("POSTAL_CODE")?.Value?.Trim();
                        string slic = pc.Element("SLIC_NR")?.Value?.Trim();
                        string loc = pc.Element("IN_BDG_LOC_NR")?.Value?.Trim();
                        if (!string.IsNullOrEmpty(code)) postcodeMap[code] = $"{slic}-{loc}";
                    }
                    if (path == null) MessageBox.Show($"Wczytano {postcodeMap.Count} kodów pocztowych.");
                    else lblStatus.Text = $"Wczytano automatycznie {postcodeMap.Count} kodów pocztowych.";
                } catch (Exception ex) { if (path == null) MessageBox.Show("Błąd XML: " + ex.Message); }
            }
        }

        private void LoadReleasedWindow()
        {
            Form f = new Form() { Text = "Zarządzanie przesyłkami RELEASED", Size = new System.Drawing.Size(600, 500), StartPosition = FormStartPosition.CenterParent };
            Label lblInfo = new Label() { Text = "Możesz wybrać wiele plików .DAT lub wkleić tekst:", Dock = DockStyle.Top, Height = 30, TextAlign = System.Drawing.ContentAlignment.BottomLeft, Padding = new Padding(5) };
            TextBox txt = new TextBox() { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical, Font = new System.Drawing.Font("Consolas", 9) };
            Panel pnlButtons = new Panel() { Dock = DockStyle.Bottom, Height = 100 };
            Button btnFile = new Button() { Text = "📁 WYBIERZ PLIKI .DAT (WIELE)", Size = new System.Drawing.Size(570, 45), Location = new System.Drawing.Point(10, 5), BackColor = System.Drawing.Color.LightCyan, FlatStyle = FlatStyle.Flat };
            Button btnProcess = new Button() { Text = "✅ DODAJ WKLEJONY TEKST", Size = new System.Drawing.Size(570, 40), Location = new System.Drawing.Point(10, 55), BackColor = System.Drawing.Color.LightSteelBlue, FlatStyle = FlatStyle.Flat };
            pnlButtons.Controls.Add(btnFile); pnlButtons.Controls.Add(btnProcess);

            Action<string[]> appendLines = (lines) => {
                int countBefore = releasedPackages.Count;
                foreach (var line in lines) {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    string[] parts = line.Split(',');
                    if (parts.Length > 5) {
                        string trackNum = parts[5].Trim();
                        if (!string.IsNullOrEmpty(trackNum)) releasedPackages.Add(trackNum);
                    }
                }
                MessageBox.Show($"Dodano nowe numery. Łącznie w pamięci: {releasedPackages.Count} (Nowych: {releasedPackages.Count - countBefore})");
            };

            btnFile.Click += (s, e) => {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Pliki DAT (*.dat)|*.dat", Multiselect = true }) {
                    if (ofd.ShowDialog() == DialogResult.OK) {
                        foreach (string file in ofd.FileNames) appendLines(File.ReadAllLines(file));
                    }
                }
            };
            btnProcess.Click += (s, e) => appendLines(txt.Lines);
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
                    if (!string.IsNullOrEmpty(num)) { mappedLocs.Add("IWMAG" + num); mappedLocs.Add("EWMAGEXP" + num); }
                } else if (isSmalls) { mappedLocs.Add("IWMSMALLS"); mappedLocs.Add("IWMSMALLS1"); mappedLocs.Add("IWMSMALLSXX"); }

                if (mappedLocs.Count > 0) {
                    for (int c = 1; c < dgv.Columns.Count; c++) {
                        string dayHeader = Regex.Match(dgv.Columns[c].HeaderText, @"\d+").Value;
                        if (int.TryParse(dayHeader, out int day)) {
                            string p1 = dgv.Rows[r].Cells[c].Value?.ToString().Trim() ?? "";
                            if (isSmalls && r + 1 < dgv.Rows.Count) {
                                string p2 = dgv.Rows[r + 1].Cells[c].Value?.ToString().Trim() ?? "";
                                if (!string.IsNullOrEmpty(p2)) p1 = string.IsNullOrEmpty(p1) ? p2 : p1 + " / " + p2;
                            }
                            if (!string.IsNullOrEmpty(p1)) foreach (var ml in mappedLocs) staffSchedule[new ScheduleKey { Loc = ml.ToUpper(), Day = day }] = p1;
                        }
                    }
                    mappedCount++; if (isSmalls) r++; 
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
                
                int colStatus = dates.Count + 2, colCity = dates.Count + 3, colZip = dates.Count + 4, colHub = dates.Count + 5, colStaff = dates.Count + 6;
                ws.Cell(1, colStatus).Value = "Status UPS"; ws.Cell(1, colCity).Value = "Lokalizacja UPS";
                ws.Cell(1, colZip).Value = "Kod Pocztowy (UPS)"; ws.Cell(1, colHub).Value = "Oddział (Mapa)";
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
                        bool isActuallyOutside = false, isReleased = releasedPackages.Contains(pkg.Key), isOutForDelivery = false;
                        bool existsAnywhereToday = pkg.Value.ContainsKey(lastDay);

                        if (!existsAnywhereToday && chkEnableUPS.Checked && !string.IsNullOrEmpty(upsLicense)) {
                            lblStatus.Text = "UPS: " + pkg.Key + "..."; Application.DoEvents();
                            var res = await GetUpsTracking(pkg.Key);
                            
                            ws.Cell(r, colStatus).Value = res.Item1;
                            ws.Cell(r, colCity).Value = res.Item2;
                            ws.Cell(r, colZip).Value = res.Item3;

                            if (!string.IsNullOrEmpty(res.Item3) && postcodeMap.TryGetValue(res.Item3, out string hub)) 
                                ws.Cell(r, colHub).Value = hub;

                            if (res.Item1.ToUpper().Contains("OUT FOR DELIVERY")) isOutForDelivery = true;
                            else if (!string.IsNullOrEmpty(res.Item2) && !res.Item2.ToUpper().Contains("STRYKOW") && !res.Item2.ToUpper().Contains("DOBRA")) isActuallyOutside = true;
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
                                    if (staffSchedule.TryGetValue(key, out string pStaff)) { cell.CreateComment().AddText(pStaff); ws.Cell(r, colStaff).Value = pStaff; }
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

        private async System.Threading.Tasks.Task<Tuple<string, string, string>> GetUpsTracking(string trackNum) {
            try {
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                string xml = "<?xml version=\"1.0\"?><AccessRequest><AccessLicenseNumber>" + upsLicense + "</AccessLicenseNumber><UserId>" + upsUser + "</UserId><Password>" + upsPass + "</Password></AccessRequest>" +
                             "<?xml version=\"1.0\"?><TrackRequest><Request><RequestAction>Track</RequestAction></Request><TrackingNumber>" + trackNum + "</TrackingNumber></TrackRequest>";
                using (var client = new HttpClient()) {
                    var resp = await client.PostAsync(upsApiUrl, new StringContent(xml, Encoding.UTF8, "application/x-www-form-urlencoded"));
                    var doc = XDocument.Parse(await resp.Content.ReadAsStringAsync());
                    
                    var shipment = doc.Descendants("Shipment").FirstOrDefault();
                    if (shipment != null) {
                        var shipTo = shipment.Element("ShipTo");
                        string zp = shipTo?.Element("Address")?.Element("PostalCode")?.Value ?? "";

                        var pkg = shipment.Element("Package");
                        var act = pkg?.Descendants("Activity").FirstOrDefault();
                        
                        string st = act?.Descendants("Status")?.FirstOrDefault()?.Descendants("StatusType")?.FirstOrDefault()?.Descendants("Description")?.FirstOrDefault()?.Value ?? "Brak";
                        string ct = act?.Descendants("ActivityLocation")?.FirstOrDefault()?.Descendants("Address")?.FirstOrDefault()?.Descendants("City")?.FirstOrDefault()?.Value ?? "Nieznane";

                        return new Tuple<string, string, string>(st, ct, zp);
                    }
                }
            } catch { }
            return new Tuple<string, string, string>("Błąd API", "---", "");
        }

        private void LoadSettings() { 
            if (File.Exists(settingsPath)) { 
                var lines = File.ReadAllLines(settingsPath); 
                if (lines.Length >= 3) { upsLicense = lines[0]; upsUser = lines[1]; upsPass = lines[2]; }
                if (lines.Length >= 4 && !string.IsNullOrWhiteSpace(lines[3])) { upsApiUrl = lines[3]; }
            } 
        }

        private void ShowSettingsWindow() {
            Form f = new Form() { Text = "Ustawienia UPS", Size = new System.Drawing.Size(350, 320), StartPosition = FormStartPosition.CenterParent };
            TextBox t1 = new TextBox() { Text = upsLicense, Dock = DockStyle.Top };
            TextBox t2 = new TextBox() { Text = upsUser, Dock = DockStyle.Top };
            TextBox t3 = new TextBox() { Text = upsPass, Dock = DockStyle.Top, UseSystemPasswordChar = true };
            TextBox t4 = new TextBox() { Text = upsApiUrl, Dock = DockStyle.Top };
            Button b = new Button() { Text = "Zapisz", Dock = DockStyle.Bottom, Height = 40 };
            
            b.Click += (s, e) => { 
                upsLicense = t1.Text; 
                upsUser = t2.Text; 
                upsPass = t3.Text; 
                if (!string.IsNullOrWhiteSpace(t4.Text)) upsApiUrl = t4.Text.Trim();
                File.WriteAllLines(settingsPath, new[] { upsLicense, upsUser, upsPass, upsApiUrl }); 
                f.Close(); 
            };

            f.Controls.Add(t4); f.Controls.Add(new Label { Text = "URL API UPS:", Dock = DockStyle.Top, Height = 25 });
            f.Controls.Add(t3); f.Controls.Add(new Label { Text = "Hasło UPS:", Dock = DockStyle.Top, Height = 25 });
            f.Controls.Add(t2); f.Controls.Add(new Label { Text = "User ID:", Dock = DockStyle.Top, Height = 25 });
            f.Controls.Add(t1); f.Controls.Add(new Label { Text = "Access License Number:", Dock = DockStyle.Top, Height = 25 });
            f.Controls.Add(b); f.ShowDialog();
        }

        [STAThread] static void Main() { Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false); Application.Run(new MainForm()); }
    }
}
