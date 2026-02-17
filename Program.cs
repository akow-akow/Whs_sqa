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

            // Ładowanie ikony z zasobów (musi być skonfigurowane w .csproj)
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream("AppIcon.ico"))
                {
                    if (stream != null) this.Icon = new System.Drawing.Icon(stream);
                }
            }
            catch { }

            // Przycisk wyboru folderu
            btnSelectFolder = new Button() { 
                Text = "📁 WYBIERZ FOLDER Z PLIKAMI AK0", 
                Dock = DockStyle.Top, 
                Height = 60,
                BackColor = System.Drawing.Color.FromArgb(230, 240, 250),
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold),
                FlatStyle = FlatStyle.Flat
            };
            btnSelectFolder.FlatAppearance.BorderColor = System.Drawing.Color.SteelBlue;
            btnSelectFolder.Click += (s, e) => SelectFolder();

            Label lbl = new Label() { 
                Text = "Lokalizacje typu 'I' znalezione w folderze:", 
                Dock = DockStyle.Top, 
                Height = 35, 
                TextAlign = System.Drawing.ContentAlignment.BottomLeft,
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Italic)
            };

            // Lista magazynów
            clbWarehouses = new CheckedListBox() { 
                Dock = DockStyle.Fill, 
                CheckOnClick = true,
                Font = new System.Drawing.Font("Segoe UI", 10),
                BorderStyle = BorderStyle.FixedSingle
            };

            // Przycisk generowania
            btnRun = new Button() { 
                Text = "🚀 GENERUJ RAPORT BRAKÓW", 
                Dock = DockStyle.Bottom, 
                Height = 70, 
                BackColor = System.Drawing.Color.FromArgb(200, 230, 201),
                Enabled = false,
                Font = new System.Drawing.Font("Segoe UI", 11, System.Drawing.FontStyle.Bold),
                FlatStyle = FlatStyle.Flat
            };
            btnRun.FlatAppearance.BorderColor = System.Drawing.Color.ForestGreen;
            btnRun.Click += BtnRun_Click;

            // Status bar
            lblStatus = new Label() { 
                Text = "Status: Oczekiwanie na folder...", 
                Dock = DockStyle.Bottom, 
                Height = 35, 
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                BackColor = System.Drawing.Color.WhiteSmoke,
                BorderStyle = BorderStyle.FixedSingle
            };

            this.Controls.Add(clbWarehouses);
            this.Controls.Add(lbl);
            this.Controls.Add(btnSelectFolder);
            this.Controls.Add(lblStatus);
            this.Controls.Add(btnRun);
        }

        private void SelectFolder()
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                // WYMUSZENIE KLASYCZNEGO DRZEWKA (zgodnie z Twoim screenem)
                fbd.AutoUpgradeEnabled = false; 
                fbd.Description = "Wybierz folder zawierający zestawienia AK0:";
                fbd.ShowNewFolderButton = true;

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
            lblStatus.Text = "Przeszukiwanie plików...";
            Application.DoEvents(); // Odświeża okno, żeby nie "zamarzło"

            var allFiles = Directory.GetFiles(selectedFolderPath, "*.xlsx");
            var validFiles = new List<(string Path, DateTime Date)>();

            foreach (var file in allFiles)
            {
                string fileName = Path.GetFileName(file);
                var match = Regex.Match(fileName, @"(\d{2}\.\d{2}\.\d{4})");
                if (match.Success && fileName.ToUpper().StartsWith("AK0"))
                {
                    if (DateTime.TryParseExact(match.Value, "dd.MM.yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime dt))
                        validFiles.Add((file, dt));
                }
            }

            sortedFiles = validFiles.OrderBy(f => f.Date).ToList();

            if (sortedFiles.Count < 2)
            {
                lblStatus.Text = "Błąd: Za mało plików!";
                MessageBox.Show("Znaleziono za mało plików AK0 (min. 2 pliki z różnymi datami).", "Błąd plików", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

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
                            if (loc.StartsWith("I", StringComparison.OrdinalIgnoreCase)) foundWarehouses.Add(loc);
                        }
                    }
                }

                foreach (var wh in foundWarehouses.OrderBy(x => x)) clbWarehouses.Items.Add(wh, true);

                lblStatus.Text = $"📁 Załadowano {sortedFiles.Count} dni. Magazynów: {foundWarehouses.Count}";
                btnRun.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Błąd odczytu Excela: " + ex.Message);
            }
        }

        private void BtnRun_Click(object sender, EventArgs e)
        {
            var selectedWarehouses = clbWarehouses.CheckedItems.Cast<string>().ToList();
            if (selectedWarehouses.Count == 0) return;

            lblStatus.Text = "⚙️ Generowanie raportu...";
            btnRun.Enabled = false;
            Application.DoEvents();

            try
            {
                string timestamp = DateTime.Now.ToString("dd.MM.yy.HH.mm.ss");
                string fileName = $"AK0_Braki_{timestamp}.xlsx";
                string reportPath = Path.Combine(selectedFolderPath, fileName);

                ProcessFinalData(selectedWarehouses, reportPath);
                
                lblStatus.Text = "✅ Gotowe!";
                MessageBox.Show($"Raport zapisany:\n{fileName}", "Sukces", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Błąd: " + ex.Message);
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
                var ws = report.Worksheets.Add("Analiza Braków");
                ws.Cell(1, 1).Value = "Package ID";
                ws.Cell(1, 1).Style.Font.Bold = true;

                for (int i = 0; i < dates.Count; i++)
                {
                    var cell = ws.Cell(1, i + 2);
                    cell.Value = dates[i].ToShortDateString();
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }
                ws.Cell(1, dates.Count + 2).Value = "Szczegóły";
                ws.Cell(1, dates.Count + 2).Style.Font.Bold = true;

                int r = 2;
                foreach (var item in allPackages)
                {
                    var firstDate = item.Value.Keys.Min();
                    // Zmiana logiki na prośbę użytkownika: pokazuje brak, jeśli paczka zniknęła w kolejnych dniach
                    var missing = dates.Where(d => d > firstDate && !item.Value.ContainsKey(d)).ToList();

                    if (missing.Any())
                    {
                        ws.Cell(r, 1).Value = item.Key;
                        for (int i = 0; i < dates.Count; i++)
                        {
                            if (item.Value.ContainsKey(dates[i])) 
                                ws.Cell(r, i + 2).Value = item.Value[dates[i]];
                            else if (dates[i] > firstDate) 
                            {
                                ws.Cell(r, i + 2).Value = "BRAK SKANU";
                                ws.Cell(r, i + 2).Style.Fill.BackgroundColor = XLColor.Salmon;
                                ws.Cell(r, i + 2).Style.Font.FontColor = XLColor.White;
                            }
                        }
                        ws.Cell(r, dates.Count + 2).Value = "Brak od: " + string.Join(", ", missing.Select(m => m.ToShortDateString()));
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
