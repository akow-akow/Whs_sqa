using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace Ak0Analyzer
{
    public class MainForm : Form
    {
        private CheckedListBox clbWarehouses;
        private Button btnRun;
        private Button btnSelectFolder;
        private Label lblStatus;
        private List<(string Path, DateTime Date)> sortedFiles;
        private string selectedFolderPath = "";

        public MainForm()
        {
            this.Text = "Warehouse Scan Quality Analysis";
            this.Size = new System.Drawing.Size(500, 650);
            this.StartPosition = FormStartPosition.CenterScreen;

            btnSelectFolder = new Button() { 
                Text = "WYBIERZ FOLDER Z PLIKAMI AK0", 
                Dock = DockStyle.Top, 
                Height = 50,
                BackColor = System.Drawing.Color.LightSkyBlue
            };
            btnSelectFolder.Click += (s, e) => SelectFolder();

            Label lbl = new Label() { 
                Text = "Wykryte lokalizacje 'I' (wybierz folder aby załadować):", 
                Dock = DockStyle.Top, 
                Height = 35, 
                TextAlign = System.Drawing.ContentAlignment.BottomLeft 
            };

            clbWarehouses = new CheckedListBox() { Dock = DockStyle.Fill, CheckOnClick = true };

            btnRun = new Button() { 
                Text = "GENERUJ RAPORT BRAKÓW", 
                Dock = DockStyle.Bottom, 
                Height = 60, 
                BackColor = System.Drawing.Color.LightGreen,
                Enabled = false 
            };
            btnRun.Click += BtnRun_Click;

            lblStatus = new Label() { 
                Text = "Oczekiwanie na wybór folderu...", 
                Dock = DockStyle.Bottom, 
                Height = 30, 
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter 
            };

            this.Controls.Add(clbWarehouses);
            this.Controls.Add(lbl);
            this.Controls.Add(btnSelectFolder);
            this.Controls.Add(lblStatus);
            this.Controls.Add(btnRun);
        }

        private void SelectFolder()
        {
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Wybierz folder zawierający zestawienia AK0";
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    selectedFolderPath = fbd.SelectedPath;
                    this.Text = $"WSQA - [{Path.GetFileName(selectedFolderPath)}]";
                    ScanFilesForLocations();
                }
            }
        }

        private void ScanFilesForLocations()
        {
            clbWarehouses.Items.Clear();
            btnRun.Enabled = false;

            var allFiles = Directory.GetFiles(selectedFolderPath, "*.xlsx");
            var validFiles = new List<(string Path, DateTime Date)>();

            foreach (var file in allFiles)
            {
                string fileName = Path.GetFileName(file);
                var match = Regex.Match(fileName, @"(\d{2}\.\d{2}\.\d{4})");
                if (match.Success && (fileName.StartsWith("AK0", StringComparison.OrdinalIgnoreCase)))
                {
                    if (DateTime.TryParseExact(match.Value, "dd.MM.yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime dt))
                    {
                        validFiles.Add((file, dt));
                    }
                }
            }

            sortedFiles = validFiles.OrderBy(f => f.Date).ToList();

            if (sortedFiles.Count < 2)
            {
                lblStatus.Text = "Błąd: Za mało plików!";
                MessageBox.Show("Wybrany folder musi zawierać co najmniej 2 pliki AK0 (format AK0_DD.MM.YYYYD.xlsx)");
                return;
            }

            lblStatus.Text = "Analizowanie nagłówków...";
            HashSet<string> foundWarehouses = new HashSet<string>();

            try
            {
                foreach (var file in sortedFiles)
                {
                    using (var workbook = new XLWorkbook(file.Path))
                    {
                        var ws = workbook.Worksheets.Contains("ak0") ? workbook.Worksheet("ak0") : workbook.Worksheet(1);
                        var rows = ws.RangeUsed().RowsUsed().Skip(1);
                        foreach (var row in rows)
                        {
                            string loc = row.Cell(1).GetString().Trim();
                            if (loc.StartsWith("I", StringComparison.OrdinalIgnoreCase))
                            {
                                foundWarehouses.Add(loc);
                            }
                        }
                    }
                }

                foreach (var wh in foundWarehouses.OrderBy(x => x))
                {
                    clbWarehouses.Items.Add(wh, true);
                }

                lblStatus.Text = $"Załadowano {sortedFiles.Count} dni. Znaleziono {foundWarehouses.Count} magazynów.";
                btnRun.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Błąd podczas odczytu plików: " + ex.Message);
            }
        }

        private void BtnRun_Click(object sender, EventArgs e)
        {
            var selectedWarehouses = clbWarehouses.CheckedItems.Cast<string>().ToList();
            if (selectedWarehouses.Count == 0)
            {
                MessageBox.Show("Musisz zaznaczyć choć jeden magazyn do analizy.");
                return;
            }

            lblStatus.Text = "Generowanie raportu Excel...";
            btnRun.Enabled = false;

            try
            {
                string reportPath = Path.Combine(selectedFolderPath, "Raport_Brakow.xlsx");
                ProcessFinalData(selectedWarehouses, reportPath);
                lblStatus.Text = "Gotowe!";
                MessageBox.Show($"Raport został zapisany w folderze:\n{reportPath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Błąd podczas generowania raportu: " + ex.Message);
            }
            finally
            {
                btnRun.Enabled = true;
                lblStatus.Text = "Gotowy";
            }
        }

        private void ProcessFinalData(List<string> activeWarehouses, string savePath)
        {
            var allPackages = new Dictionary<string, SortedDictionary<DateTime, string>>();
            var dates = sortedFiles.Select(f => f.Date).ToList();

            foreach (var file in sortedFiles)
            {
                using (var workbook = new XLWorkbook(file.Path))
                {
                    var ws = workbook.Worksheets.Contains("ak0") ? workbook.Worksheet("ak0") : workbook.Worksheet(1);
                    var rows = ws.RangeUsed().RowsUsed().Skip(1);

                    foreach (var row in rows)
                    {
                        string loc = row.Cell(1).GetString().Trim();
                        string pkg = row.Cell(2).GetString().Trim();

                        if (activeWarehouses.Contains(loc))
                        {
                            if (!allPackages.ContainsKey(pkg)) allPackages[pkg] = new SortedDictionary<DateTime, string>();
                            allPackages[pkg][file.Date] = loc;
                        }
                    }
                }
            }

            using (var report = new XLWorkbook())
            {
                var ws = report.Worksheets.Add("Analiza");
                ws.Cell(1, 1).Value = "Package ID";
                for (int i = 0; i < dates.Count; i++)
                {
                    var cell = ws.Cell(1, i + 2);
                    cell.Value = dates[i].ToShortDateString();
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                }
                ws.Cell(1, dates.Count + 2).Value = "Szczegóły";

                int r = 2;
                foreach (var item in allPackages)
                {
                    var first = item.Value.Keys.Min();
                    var last = item.Value.Keys.Max();
                    var missing = dates.Where(d => d > first && d < last && !item.Value.ContainsKey(d)).ToList();

                    if (missing.Any())
                    {
                        ws.Cell(r, 1).Value = item.Key;
                        for (int i = 0; i < dates.Count; i++)
                        {
                            if (item.Value.ContainsKey(dates[i])) 
                                ws.Cell(r, i + 2).Value = item.Value[dates[i]];
                            else if (dates[i] > first && dates[i] < last) 
                            {
                                ws.Cell(r, i + 2).Value = "BRAK SKANU";
                                ws.Cell(r, i + 2).Style.Fill.BackgroundColor = XLColor.Red;
                                ws.Cell(r, i + 2).Style.Font.FontColor = XLColor.White;
                            }
                        }
                        ws.Cell(r, dates.Count + 2).Value = "Brak skanu: " + string.Join(", ", missing.Select(m => m.ToShortDateString()));
                        r++;
                    }
                }

                ws.Columns().AdjustToContents();
                report.SaveAs(savePath);
            }
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
