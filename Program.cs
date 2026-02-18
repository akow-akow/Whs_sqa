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
            this.Text = "WSQA PRO - Analiza AK0";
            this.Size = new System.Drawing.Size(550, 750);
            this.StartPosition = FormStartPosition.CenterScreen;

            try {
                var assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream("AppIcon.ico"))
                    if (stream != null) this.Icon = new System.Drawing.Icon(stream);
            } catch { }

            FlowLayoutPanel topPanel = new FlowLayoutPanel() { Dock = DockStyle.Top, Height = 140, Padding = new Padding(10) };

            btnSelectFolder = new Button() { Text = "📁 1. WYBIERZ FOLDER AK0", Size = new System.Drawing.Size(245, 60), BackColor = System.Drawing.Color.LightSkyBlue, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btnSelectFolder.Click += (s, e) => SelectFolder();

            btnLoadSchedule = new Button() { Text = "📅 2. WCZYTAJ GRAFIK (ARKUSZ)", Size = new System.Drawing.Size(245, 60), BackColor = System.Drawing.Color.NavajoWhite, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btnLoadSchedule.Click += (s, e) => LoadScheduleWindow();

            topPanel.Controls.Add(btnSelectFolder);
            topPanel.Controls.Add(btnLoadSchedule);

            clbWarehouses = new CheckedListBox() { Dock = DockStyle.Fill, CheckOnClick = true, Font = new System.Drawing.Font("Segoe UI", 10), BorderStyle = BorderStyle.FixedSingle };
            
            btnRun = new Button() { Text = "🚀 3. GENERUJ RAPORT", Dock = DockStyle.Bottom, Height = 70, BackColor = System.Drawing.Color.LightGreen, Enabled = false, Font = new System.Drawing.Font("Segoe UI", 11, System.Drawing.FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btnRun.Click += BtnRun_Click;

            lblStatus = new Label() { Text = "Gotowy", Dock = DockStyle.Bottom, Height = 40, TextAlign = System.Drawing.ContentAlignment.MiddleCenter, BackColor = System.Drawing.Color.WhiteSmoke, BorderStyle = BorderStyle.FixedSingle };

            this.Controls.Add(clbWarehouses);
            this.Controls.Add(new Label() { Text = " Magazyny do analizy:", Dock = DockStyle.Top, Height = 25, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold) });
            this.Controls.Add(topPanel);
            this.Controls.Add(lblStatus);
            this.Controls.Add(btnRun);
        }

        private void LoadScheduleWindow()
        {
            Form f = new Form() { Text = "Arkusz Grafiku - Wklej dane (Ctrl+V)", Size = new System.Drawing.Size(800, 500), StartPosition = FormStartPosition.CenterParent };
            DataGridView dgv = new DataGridView() { 
                Dock = DockStyle.Fill, 
                BackgroundColor = System.Drawing.Color.White,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                AllowUserToAddRows = false
            };
            
            Button btnProcess = new Button() { Text = "ZATWIERDŹ DANE Z TABELI", Dock = DockStyle.Bottom, Height = 50, BackColor = System.Drawing.Color.PaleGreen };

            // Obsługa skrótu Ctrl+V w tabeli
            dgv.KeyDown += (s, e) => {
                if (e.Control && e.KeyCode == Keys.V) {
                    PasteToDgv(dgv);
                }
            };

            btnProcess.Click += (s, e) => {
                ProcessDgvData(dgv);
                f.Close();
            };

            f.Controls.Add(dgv);
            f.Controls.Add(new Label() { Text = "Kliknij w tabelę i wklej dane z Excela (Ctrl+V). Pierwszy wiersz musi być numerami dni.", Dock = DockStyle.Top, Height = 30 });
            f.Controls.Add(btnProcess);
            f.ShowDialog();
        }

        private void PasteToDgv(DataGridView dgv)
        {
            string clipboardText = Clipboard.GetText();
            if (string.IsNullOrEmpty(clipboardText)) return;

            dgv.Rows.Clear();
            dgv.Columns.Clear();

            string[] lines = clipboardText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0) return;

            // Tworzenie kolumn na podstawie pierwszego wiersza
            string[] firstLineCells = lines[0].Split('\t');
            for (int i = 0; i < firstLineCells.Length; i++) {
                dgv.Columns.Add("col" + i, firstLineCells[i]);
            }

            // Dodawanie reszty wierszy
            for (int i = 1; i < lines.Length; i++) {
                dgv.Rows.Add(lines[i].Split('\t'));
            }
        }

        private void ProcessDgvData(DataGridView dgv)
        {
            staffSchedule.Clear();
            if (dgv.Columns.Count < 2) return;

            // Nagłówki kolumn to nasze dni miesiąca
            for (int r = 0; r < dgv.Rows.Count; r++) {
                string rawLoc = dgv.Rows[r].Cells[0].Value?.ToString() ?? "";
                string mappedLoc = MapLocationName(rawLoc.ToLower().Trim());

                for (int c = 1; c < dgv.Columns.Count; c++) {
                    string dayHeader = dgv.Columns[c].HeaderText;
                    if (int.TryParse(dayHeader, out int dayNum)) {
                        string person = dgv.Rows[r].Cells[c].Value?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(person)) {
                            staffSchedule[(mappedLoc, dayNum)] = person;
                            if (mappedLoc == "IWMSMALLS1") staffSchedule[("IWMSMALLSXX", dayNum)] = person;
                        }
                    }
                }
            }
            lblStatus.Text = "Grafik przetworzony pomyślnie!";
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
                if (fbd.ShowDialog() == DialogResult.OK) {
                    selectedFolderPath = fbd.SelectedPath;
                    ScanFiles();
                }
            }
        }

        private void ScanFiles()
        {
            clbWarehouses.Items.Clear();
            btnRun.Enabled = false;
            if (string.IsNullOrEmpty(selectedFolderPath)) return;

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
            if (sortedFiles.Count < 2) { lblStatus.Text = "Min. 2 pliki AK0!"; return; }

            HashSet<string> locs = new HashSet<string>();
            try {
                foreach (var f in sortedFiles) {
                    using (var wb = new XLWorkbook(f.Item1)) {
                        var ws = wb.Worksheets.FirstOrDefault(w => w.Name.ToLower() == "ak0") ?? wb.Worksheets.FirstOrDefault();
                        if (ws == null) continue;
                        foreach (var row in ws.RangeUsed().RowsUsed().Skip(1)) {
                            string l = row.Cell(1).GetString().Trim();
                            if (l.StartsWith("I", StringComparison.OrdinalIgnoreCase)) locs.Add(l);
                        }
                    }
                }
                foreach (var l in locs.OrderBy(x => x)) clbWarehouses.Items.Add(l, false);
                btnRun.Enabled = true;
                lblStatus.Text = $"Wczytano {sortedFiles.Count} dni.";
            } catch (Exception ex) { MessageBox.Show("Błąd: " + ex.Message); }
        }

        private void BtnRun_Click(object sender, EventArgs e)
        {
            btnRun.Enabled = false;
            lblStatus.Text = "Generowanie...";
            Application.DoEvents();
            try { GenerateReport(); lblStatus.Text = "Gotowe!"; }
            catch (Exception ex) { MessageBox.Show("Błąd: " + ex.Message); }
            finally { btnRun.Enabled = true; }
        }

        private void GenerateReport()
        {
            var selectedLocs = clbWarehouses.CheckedItems.Cast<string>().ToList();
            var data = new Dictionary<string, SortedDictionary<DateTime, string>>();
            var dates = sortedFiles.Select(x => x.Item2).ToList();

            foreach (var f in sortedFiles) {
                using (var wb = new XLWorkbook(f.Item1)) {
                    var ws = wb.Worksheets.FirstOrDefault(w => w.Name.ToLower() == "ak0") ?? wb.Worksheets.FirstOrDefault();
                    foreach (var row in ws.RangeUsed().RowsUsed().Skip(1)) {
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
                var userStats = new Dictionary<string, int>();

                ws.Cell(1, 1).Value = "Package ID";
                for (int i = 0; i < dates.Count; i++) ws.Cell(1, i + 2).Value = dates[i].ToShortDateString();

                int r = 2;
                foreach (var pkg in data) {
                    var first = pkg.Value.Keys.Min();
                    var last = pkg.Value.Keys.Max();
                    bool hasError = false;
                    for (int i = 0; i < dates.Count; i++) {
                        if (dates[i] > first && dates[i] < last && !pkg.Value.ContainsKey(dates[i])) {
                            var next = pkg.Value.Keys.Where(d => d > dates[i]).Min();
                            if ((next - dates[i]).TotalDays <= 3) { hasError = true; break; }
                        }
                    }

                    if (hasError) {
                        ws.Cell(r, 1).Value = pkg.Key;
                        for (int i = 0; i < dates.Count; i++) {
                            DateTime d = dates[i];
                            if (pkg.Value.ContainsKey(d)) ws.Cell(r, i + 2).Value = pkg.Value[d];
                            else if (d > first && d < last) {
                                var next = pkg.Value.Keys.Where(dt => dt > d).Min();
                                if ((next - d).TotalDays <= 3) {
                                    var cell = ws.Cell(r, i + 2);
                                    cell.Value = "BRAK SKANU";
                                    cell.Style.Fill.BackgroundColor = XLColor.Salmon;
                                    string loc = pkg.Value[first];
                                    if (staffSchedule.TryGetValue((loc, d.Day), out string person)) {
                                        cell.CreateComment().AddText(person);
                                        if (!userStats.ContainsKey(person)) userStats[person] = 0;
                                        userStats[person]++;
                                    }
                                }
                            }
                        }
                        r++;
                    }
                }

                wsStat.Cell(1, 1).Value = "User"; wsStat.Cell(1, 2).Value = "Braki";
                int sr = 2;
                foreach (var s in userStats.OrderByDescending(x => x.Value)) {
                    wsStat.Cell(sr, 1).Value = s.Key; wsStat.Cell(sr, 2).Value = s.Value; sr++;
                }
                ws.Columns().AdjustToContents();
                wsStat.Columns().AdjustToContents();
                report.SaveAs(Path.Combine(selectedFolderPath, $"Raport_{DateTime.Now:ddMMyy_HHmm}.xlsx"));
            }
        }

        [STAThread] static void Main() { Application.EnableVisualStyles(); Application.Run(new MainForm()); }
    }
}
