using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace Ak0Analyzer
{
    public class MainForm : Form
    {
        private CheckedListBox clbWarehouses;
        private Button btnRun, btnSelectFolder, btnLoadSchedule;
        private Label lblStatus;
        private List<(string Path, DateTime Date)> sortedFiles;
        private string selectedFolderPath = "";
        private Dictionary<(string Loc, int Day), string> staffSchedule = new Dictionary<(string, int), string>();

        public MainForm()
        {
            this.Text = "Warehouse Scan Quality Analysis PRO";
            this.Size = new System.Drawing.Size(550, 750);
            this.StartPosition = FormStartPosition.CenterScreen;

            try {
                var assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream("AppIcon.ico"))
                    if (stream != null) this.Icon = new System.Drawing.Icon(stream);
            } catch { }

            // Panel przycisków górnych
            FlowLayoutPanel topPanel = new FlowLayoutPanel() { Dock = DockStyle.Top, Height = 130 };

            btnSelectFolder = new Button() { Text = "📁 1. WYBIERZ FOLDER AK0", Size = new System.Drawing.Size(260, 60), BackColor = System.Drawing.Color.LightSkyBlue, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btnSelectFolder.Click += (s, e) => SelectFolder();

            btnLoadSchedule = new Button() { Text = "📅 2. WCZYTAJ GRAFIK", Size = new System.Drawing.Size(260, 60), BackColor = System.Drawing.Color.NavajoWhite, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btnLoadSchedule.Click += (s, e) => LoadScheduleWindow();

            topPanel.Controls.Add(btnSelectFolder);
            topPanel.Controls.Add(btnLoadSchedule);

            clbWarehouses = new CheckedListBox() { Dock = DockStyle.Fill, CheckOnClick = true, Font = new System.Drawing.Font("Segoe UI", 10), BorderStyle = BorderStyle.FixedSingle };
            btnRun = new Button() { Text = "🚀 3. GENERUJ RAPORT Z GRAFIKIEM", Dock = DockStyle.Bottom, Height = 70, BackColor = System.Drawing.Color.LightGreen, Enabled = false, Font = new System.Drawing.Font("Segoe UI", 11, System.Drawing.FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btnRun.Click += BtnRun_Click;

            lblStatus = new Label() { Text = "Wybierz folder i wgraj grafik...", Dock = DockStyle.Bottom, Height = 40, TextAlign = System.Drawing.ContentAlignment.MiddleCenter, BackColor = System.Drawing.Color.WhiteSmoke, BorderStyle = BorderStyle.FixedSingle };

            this.Controls.Add(clbWarehouses);
            this.Controls.Add(new Label() { Text = "Magazyny (Lokalizacje 'I'):", Dock = DockStyle.Top, Height = 25 });
            this.Controls.Add(topPanel);
            this.Controls.Add(lblStatus);
            this.Controls.Add(btnRun);
        }

        private void LoadScheduleWindow()
        {
            Form f = new Form() { Text = "Wklej dane grafiku (Ctrl+V)", Size = new System.Drawing.Size(600, 400), StartPosition = FormStartPosition.CenterParent };
            TextBox tb = new TextBox() { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Both };
            Button btnSave = new Button() { Text = "ZAPISZ GRAFIK", Dock = DockStyle.Bottom, Height = 50, BackColor = System.Drawing.Color.PaleGreen };
            
            btnSave.Click += (s, e) => {
                ParseSchedule(tb.Text);
                f.Close();
            };
            f.Controls.Add(tb);
            f.Controls.Add(btnSave);
            f.ShowDialog();
        }

        private void ParseSchedule(string data)
        {
            staffSchedule.Clear();
            var lines = data.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length < 2) return;

            var daysHeader = lines[0].Split('\t'); // Dni miesiąca (A1, B1...)

            for (int i = 1; i < lines.Length; i++)
            {
                var cols = lines[i].Split('\t');
                if (cols.Length < 2) continue;

                string rawLoc = cols[0].ToLower();
                string mappedLoc = MapLocationName(rawLoc);

                for (int dayIdx = 1; dayIdx < cols.Length; dayIdx++)
                {
                    if (int.TryParse(daysHeader[dayIdx], out int dayNum))
                    {
                        string person = cols[dayIdx].Trim();
                        if (!string.IsNullOrEmpty(person))
                        {
                            staffSchedule[(mappedLoc, dayNum)] = person;
                            // Obsługa Smalls dla obu magazynów
                            if (mappedLoc == "IWMSMALLS1") staffSchedule[("IWMSMALLSXX", dayNum)] = person;
                        }
                    }
                }
            }
            lblStatus.Text = "Grafik wczytany poprawnie!";
        }

        private string MapLocationName(string raw)
        {
            if (raw.Contains("100")) return "IWMAG100";
            if (raw.Contains("mag")) {
                var num = Regex.Match(raw, @"\d+").Value;
                return "IWMAGAZYN" + num;
            }
            if (raw.Contains("smalls")) return "IWMSMALLS1";
            return raw.ToUpper();
        }

        private void SelectFolder()
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog() { AutoUpgradeEnabled = false })
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    selectedFolderPath = fbd.SelectedPath;
                    ScanFiles();
                }
            }
        }

        private void ScanFiles()
        {
            clbWarehouses.Items.Clear();
            var files = Directory.GetFiles(selectedFolderPath, "*.xlsx");
            var valid = new List<(string, DateTime)>();
            foreach (var f in files)
            {
                var m = Regex.Match(Path.GetFileName(f), @"(\d{2}\.\d{2}\.\d{4})");
                if (m.Success && Path.GetFileName(f).ToUpper().StartsWith("AK0"))
                    if (DateTime.TryParseExact(m.Value, "dd.MM.yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime dt))
                        valid.Add((f, dt));
            }
            sortedFiles = valid.OrderBy(x => x.Item2).ToList();
            
            if (sortedFiles.Count < 2) { MessageBox.Show("Min. 2 pliki AK0!"); return; }

            HashSet<string> locs = new HashSet<string>();
            foreach (var f in sortedFiles)
            {
                using (var wb = new XLWorkbook(f.Item1))
                {
                    var ws = wb.Worksheets.First();
                    foreach (var row in ws.RangeUsed().RowsUsed().Skip(1))
                    {
                        string l = row.Cell(1).GetString().Trim();
                        if (l.StartsWith("I", StringComparison.OrdinalIgnoreCase)) locs.Add(l);
                    }
                }
            }
            foreach (var l in locs.OrderBy(x => x)) clbWarehouses.Items.Add(l, true);
            btnRun.Enabled = true;
            lblStatus.Text = $"Wczytano {sortedFiles.Count} dni.";
        }

        private void BtnRun_Click(object sender, EventArgs e)
        {
            btnRun.Enabled = false;
            lblStatus.Text = "Przetwarzanie danych...";
            Application.DoEvents();

            try {
                GenerateReport();
                MessageBox.Show("Raport wygenerowany!");
            } catch (Exception ex) { MessageBox.Show("Błąd: " + ex.Message); }
            finally { btnRun.Enabled = true; lblStatus.Text = "Gotowy"; }
        }

        private void GenerateReport()
        {
            var selectedLocs = clbWarehouses.CheckedItems.Cast<string>().ToList();
            var data = new Dictionary<string, SortedDictionary<DateTime, string>>();
            var dates = sortedFiles.Select(x => x.Item2).ToList();

            foreach (var f in sortedFiles)
            {
                using (var wb = new XLWorkbook(f.Item1))
                {
                    var ws = wb.Worksheets.First();
                    foreach (var row in ws.RangeUsed().RowsUsed().Skip(1))
                    {
                        string l = row.Cell(1).GetString().Trim();
                        string p = row.Cell(2).GetString().Trim();
                        if (selectedLocs.Contains(l))
                        {
                            if (!data.ContainsKey(p)) data[p] = new SortedDictionary<DateTime, string>();
                            data[p][f.Item2] = l;
                        }
                    }
                }
            }

            using (var report = new XLWorkbook())
            {
                var ws = report.Worksheets.Add("Analiza");
                var wsStat = report.Worksheets.Add("Podsumowanie Userów");
                var userStats = new Dictionary<string, int>();

                // Nagłówki
                ws.Cell(1, 1).Value = "Package ID";
                for (int i = 0; i < dates.Count; i++) ws.Cell(1, i + 2).Value = dates[i].ToShortDateString();

                int r = 2;
                foreach (var pkg in data)
                {
                    var first = pkg.Value.Keys.Min();
                    var last = pkg.Value.Keys.Max();
                    
                    // Logika: Braki tylko jeśli przerwa <= 3 dni. Jeśli więcej - uznajemy za wyjazd.
                    bool hasError = false;
                    for (int i = 0; i < dates.Count - 1; i++)
                    {
                        if (dates[i] >= first && dates[i] < last && !pkg.Value.ContainsKey(dates[i]))
                        {
                            var nextScan = pkg.Value.Keys.Where(d => d > dates[i]).Min();
                            if ((nextScan - dates[i]).TotalDays <= 3) { hasError = true; break; }
                        }
                    }

                    if (hasError)
                    {
                        ws.Cell(r, 1).Value = pkg.Key;
                        for (int i = 0; i < dates.Count; i++)
                        {
                            DateTime currentD = dates[i];
                            if (pkg.Value.ContainsKey(currentD))
                                ws.Cell(r, i + 2).Value = pkg.Value[currentD];
                            else if (currentD > first && currentD < last)
                            {
                                var nextScan = pkg.Value.Keys.Where(d => d > currentD).Min();
                                if ((nextScan - currentD).TotalDays <= 3)
                                {
                                    var cell = ws.Cell(r, i + 2);
                                    cell.Value = "BRAK SKANU";
                                    cell.Style.Fill.BackgroundColor = XLColor.Salmon;
                                    
                                    // Pobieranie osoby z grafiku
                                    string locAtError = pkg.Value[first]; // Zakładamy lokalizację z pierwszego skanu
                                    if (staffSchedule.TryGetValue((locAtError, currentD.Day), out string person))
                                    {
                                        cell.CreateComment().AddText("Odpowiedzialny: " + person);
                                        if (!userStats.ContainsKey(person)) userStats[person] = 0;
                                        userStats[person]++;
                                    }
                                }
                            }
                        }
                        r++;
                    }
                }

                // Arkusz statystyk
                wsStat.Cell(1, 1).Value = "Pracownik";
                wsStat.Cell(1, 2).Value = "Liczba Braków Skany";
                int sr = 2;
                foreach (var stat in userStats.OrderByDescending(x => x.Value))
                {
                    wsStat.Cell(sr, 1).Value = stat.Key;
                    wsStat.Cell(sr, 2).Value = stat.Value;
                    sr++;
                }

                ws.Columns().AdjustToContents();
                wsStat.Columns().AdjustToContents();
                report.SaveAs(Path.Combine(selectedFolderPath, $"Raport_PRO_{DateTime.Now:dd.MM.yy.HH.mm}.xlsx"));
            }
        }

        [STAThread] static void Main() { Application.EnableVisualStyles(); Application.Run(new MainForm()); }
    }
}
