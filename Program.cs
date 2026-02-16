using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions; // Dodane dla lepszego rozpoznawania nazw
using System.Windows.Forms;
using ClosedXML.Excel;

namespace Ak0Analyzer
{
    public class MainForm : Form
    {
        private CheckedListBox clbWarehouses;
        private Button btnRun;
        private Label lblStatus;
        private string[] initialWarehouses = {
            "IWMAG100", "IWMAGAZYN", "IWMAGAZYN1", "IWMAGAZYN10", "IWMAGAZYN2", "IWMAGAZYN3",
            "IWMAGAZYN4", "IWMAGAZYN5", "IWMAGAZYN6", "IWMAGAZYN7", "IWMAGAZYN8", "IWMAGAZYN9",
            "IWMAGCEL1", "IWMAGCEL10", "IWMAGCEL12", "IWMAGCEL13", "IWMAGCEL16", "IWMAGCEL23",
            "IWMAGCEL6", "IWMAGCEL7", "IWMAGHRTS", "IWMAGTZR", "IWMSMALLS1", "IWMSMALLSXX"
        };

        public MainForm()
        {
            this.Text = "Ak0 Analyzer - v1.1";
            this.Size = new System.Drawing.Size(400, 550);
            this.StartPosition = FormStartPosition.CenterScreen;

            Label lbl = new Label() { Text = "Magazyny do sprawdzenia:", Dock = DockStyle.Top, Height = 30, TextAlign = System.Drawing.ContentAlignment.BottomLeft };
            clbWarehouses = new CheckedListBox() { Dock = DockStyle.Fill, CheckOnClick = true };
            clbWarehouses.Items.AddRange(initialWarehouses);
            
            for (int i = 0; i < clbWarehouses.Items.Count; i++) clbWarehouses.SetItemChecked(i, true);

            btnRun = new Button() { Text = "ANALIZUJ PLIKI AK0", Dock = DockStyle.Bottom, Height = 60, BackColor = System.Drawing.Color.LightBlue };
            btnRun.Click += BtnRun_Click;

            lblStatus = new Label() { Text = "Gotowy", Dock = DockStyle.Bottom, Height = 30, TextAlign = System.Drawing.ContentAlignment.MiddleCenter };

            this.Controls.Add(clbWarehouses);
            this.Controls.Add(lbl);
            this.Controls.Add(lblStatus);
            this.Controls.Add(btnRun);
        }

        private void BtnRun_Click(object sender, EventArgs e)
        {
            var selectedWarehouses = clbWarehouses.CheckedItems.Cast<string>().ToList();
            
            // Pobieramy pliki Ak0 wspierając spację i podkreślnik
            var allFiles = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.xlsx");
            var validFiles = new List<(string Path, DateTime Date)>();

            foreach (var file in allFiles)
            {
                string fileName = Path.GetFileName(file);
                // Regex szuka wzoru DD.MM.YYYY niezależnie czy przed nim jest spacja czy _
                var match = Regex.Match(fileName, @"(\d{2}\.\d{2}\.\d{4})");
                if (match.Success && (fileName.StartsWith("AK0_", StringComparison.OrdinalIgnoreCase) || 
                                     fileName.StartsWith("AK0 ", StringComparison.OrdinalIgnoreCase)))
                {
                    if (DateTime.TryParseExact(match.Value, "dd.MM.yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime dt))
                    {
                        validFiles.Add((file, dt));
                    }
                }
            }

            var sortedFiles = validFiles.OrderBy(f => f.Date).ToList();

            if (sortedFiles.Count < 2) {
                MessageBox.Show("Znaleziono za mało pasujących plików (wymagane min. 2).\n\nWymagany format: AK0_DD.MM.YYYYD.xlsx lub AK0 DD.MM.YYYYD.xlsx");
                return;
            }

            lblStatus.Text = "Przetwarzanie...";
            btnRun.Enabled = false;

            try {
                ProcessData(sortedFiles, selectedWarehouses);
                lblStatus.Text = "Sukces!";
                MessageBox.Show($"Analiza zakończona.\nPrzeanalizowano pliki z {sortedFiles.Count} dni.");
            }
            catch (Exception ex) {
                MessageBox.Show("Wystąpił błąd podczas odczytu danych: " + ex.Message);
            }
            finally {
                btnRun.Enabled = true;
                lblStatus.Text = "Gotowy";
            }
        }

        private void ProcessData(List<(string Path, DateTime Date)> files, List<string> activeWarehouses)
        {
            var allPackages = new Dictionary<string, SortedDictionary<DateTime, string>>();
            var dates = files.Select(f => f.Date).ToList();

            foreach (var file in files)
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
                var ws = report.Worksheets.Add("Braki");
                
                // Nagłówki
                ws.Cell(1, 1).Value = "Package ID";
                for (int i = 0; i < dates.Count; i++)
                {
                    var cell = ws.Cell(1, i + 2);
                    cell.Value = dates[i].ToShortDateString();
                    cell.Style.Font.Bold = true;
                }
                ws.Cell(1, dates.Count + 2).Value = "Podsumowanie";
                ws.Cell(1, dates.Count + 2).Style.Font.Bold = true;

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
                            {
                                ws.Cell(r, i + 2).Value = item.Value[dates[i]];
                            }
                            else if (dates[i] > first && dates[i] < last) 
                            {
                                ws.Cell(r, i + 2).Value = "BRAK SKANU";
                                ws.Cell(r, i + 2).Style.Fill.BackgroundColor = XLColor.Red;
                                ws.Cell(r, i + 2).Style.Font.FontColor = XLColor.White;
                            }
                        }
                        ws.Cell(r, dates.Count + 2).Value = "Luki: " + string.Join(", ", missing.Select(m => m.ToShortDateString()));
                        r++;
                    }
                }

                ws.Columns().AdjustToContents();
                report.SaveAs("Raport_Brakow.xlsx");
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
