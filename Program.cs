using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
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
            this.Text = "Ak0 Analyzer - Wybór Magazynów";
            this.Size = new System.Drawing.Size(400, 550);

            Label lbl = new Label() { Text = "Zaznacz magazyny do SPRAWDZENIA:", Dock = DockStyle.Top, Height = 30 };
            clbWarehouses = new CheckedListBox() { Dock = DockStyle.Fill, CheckOnClick = true };
            clbWarehouses.Items.AddRange(initialWarehouses);
            
            // Domyślnie zaznacz wszystko
            for (int i = 0; i < clbWarehouses.Items.Count; i++) clbWarehouses.SetItemChecked(i, true);

            btnRun = new Button() { Text = "Analizuj pliki Ak0 w folderze", Dock = DockStyle.Bottom, Height = 50 };
            btnRun.Click += BtnRun_Click;

            lblStatus = new Label() { Text = "Gotowy", Dock = DockStyle.Bottom, Height = 30 };

            this.Controls.Add(clbWarehouses);
            this.Controls.Add(lbl);
            this.Controls.Add(lblStatus);
            this.Controls.Add(btnRun);
        }

        private void BtnRun_Click(object sender, EventArgs e)
        {
            var selectedWarehouses = clbWarehouses.CheckedItems.Cast<string>().ToList();
            var files = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "Ak0_*.xlsx")
                                 .OrderBy(f => f).ToList();

            if (files.Count < 2) {
                MessageBox.Show("W folderze muszą być min. 2 pliki Ak0_DD.MM.RRRRD.xlsx");
                return;
            }

            lblStatus.Text = "Przetwarzanie...";
            btnRun.Enabled = false;

            try {
                ProcessData(files, selectedWarehouses);
                lblStatus.Text = "Zakończono! Sprawdź Raport_Brakow.xlsx";
                MessageBox.Show("Raport wygenerowany pomyślnie.");
            }
            catch (Exception ex) {
                MessageBox.Show("Błąd: " + ex.Message);
            }
            finally {
                btnRun.Enabled = true;
            }
        }

        private void ProcessData(List<string> files, List<string> activeWarehouses)
        {
            var allPackages = new Dictionary<string, SortedDictionary<DateTime, string>>();
            var dates = new List<DateTime>();

            foreach (var file in files)
            {
                DateTime fileDate = DateTime.ParseExact(Path.GetFileName(file).Split('_', 'D')[1], "dd.MM.yyyy", null);
                dates.Add(fileDate);

                using (var workbook = new XLWorkbook(file))
                {
                    var ws = workbook.Worksheets.Contains("ak0") ? workbook.Worksheet("ak0") : workbook.Worksheet(1);
                    var rows = ws.RangeUsed().RowsUsed().Skip(1);

                    foreach (var row in rows)
                    {
                        string loc = row.Cell(1).GetString();
                        string pkg = row.Cell(2).GetString();

                        if (activeWarehouses.Contains(loc))
                        {
                            if (!allPackages.ContainsKey(pkg)) allPackages[pkg] = new SortedDictionary<DateTime, string>();
                            allPackages[pkg][fileDate] = loc;
                        }
                    }
                }
            }

            using (var report = new XLWorkbook())
            {
                var ws = report.Worksheets.Add("Braki");
                ws.Cell(1, 1).Value = "Package ID";
                for (int i = 0; i < dates.Count; i++) ws.Cell(1, i + 2).Value = dates[i].ToShortDateString();
                ws.Cell(1, dates.Count + 2).Value = "Podsumowanie";

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
                            if (item.Value.ContainsKey(dates[i])) ws.Cell(r, i + 2).Value = item.Value[dates[i]];
                            else if (dates[i] > first && dates[i] < last) {
                                ws.Cell(r, i + 2).Value = "BRAK SKANU";
                                ws.Cell(r, i + 2).Style.Fill.BackgroundColor = XLColor.Red;
                            }
                        }
                        ws.Cell(r, dates.Count + 2).Value = "Brak skanu w dniach: " + string.Join(", ", missing.Select(m => m.ToShortDateString()));
                        r++;
                    }
                }
                report.SaveAs("Raport_Brakow.xlsx");
            }
        }

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.Run(new MainForm());
        }
    }
}
