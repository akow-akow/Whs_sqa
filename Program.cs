private void ScanFiles()
        {
            clbWarehouses.Items.Clear();
            btnRun.Enabled = false;
            
            if (string.IsNullOrEmpty(selectedFolderPath) || !Directory.Exists(selectedFolderPath))
                return;

            var files = Directory.GetFiles(selectedFolderPath, "*.xlsx");
            var valid = new List<(string, DateTime)>();

            foreach (var f in files)
            {
                var fileName = Path.GetFileName(f);
                // Ignorujemy pliki tymczasowe Excela (zaczynające się od ~$)
                if (fileName.StartsWith("~$")) continue;

                var m = Regex.Match(fileName, @"(\d{2}\.\d{2}\.\d{4})");
                if (m.Success && fileName.ToUpper().StartsWith("AK0"))
                {
                    if (DateTime.TryParseExact(m.Value, "dd.MM.yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime dt))
                        valid.Add((f, dt));
                }
            }

            sortedFiles = valid.OrderBy(x => x.Item2).ToList();
            
            if (sortedFiles == null || sortedFiles.Count < 2) 
            { 
                lblStatus.Text = "Błąd: Potrzebne min. 2 pliki AK0!";
                return; 
            }

            HashSet<string> locs = new HashSet<string>();
            try
            {
                foreach (var f in sortedFiles)
                {
                    using (var wb = new XLWorkbook(f.Item1))
                    {
                        // Sprawdzamy czy plik ma arkusze
                        if (wb.Worksheets.Count == 0) continue;

                        // Próbujemy znaleźć arkusz "ak0" lub bierzemy pierwszy dostępny
                        var ws = wb.Worksheets.FirstOrDefault(w => w.Name.ToLower() == "ak0") ?? wb.Worksheets.First();
                        
                        var rows = ws.RangeUsed()?.RowsUsed();
                        if (rows == null) continue;

                        foreach (var row in rows.Skip(1))
                        {
                            var cell = row.Cell(1);
                            if (cell == null || cell.IsEmpty()) continue;

                            string l = cell.GetString().Trim();
                            if (l.StartsWith("I", StringComparison.OrdinalIgnoreCase)) 
                                locs.Add(l);
                        }
                    }
                }

                if (locs.Count > 0)
                {
                    foreach (var l in locs.OrderBy(x => x)) 
                        clbWarehouses.Items.Add(l, false);
                    
                    btnRun.Enabled = true;
                    lblStatus.Text = $"Wczytano {sortedFiles.Count} dni. Wybierz magazyny.";
                }
                else
                {
                    lblStatus.Text = "Nie znaleziono lokalizacji zaczynających się na 'I'.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Błąd podczas skanowania plików: {ex.Message}\nUpewnij się, że pliki Excel nie są otwarte w innym programie.", "Błąd odczytu");
                lblStatus.Text = "Błąd krytyczny plików.";
            }
        }
