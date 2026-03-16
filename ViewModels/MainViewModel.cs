using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using ReportGenerator.Helpers;
using ReportGenerator.Services;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input; 
using Colors = QuestPDF.Helpers.Colors;

namespace ReportGenerator.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly IDatabaseService _db;

        public ObservableCollection<string> Tables { get; } = new ObservableCollection<string>();
        public ObservableCollection<ColumnItem> Columns { get; } = new ObservableCollection<ColumnItem>();
        public ObservableCollection<string> ExportFormats { get; } = new ObservableCollection<string> { "Excel", "CSV", "PDF" };
        public ObservableCollection<int> Hours { get; } = new ObservableCollection<int>(Enumerable.Range(0, 24));
        public ObservableCollection<int> Minutes { get; } = new ObservableCollection<int>(Enumerable.Range(0, 60));

        private string _selectedExportFormat = "Excel";
        public string SelectedExportFormat
        {
            get => _selectedExportFormat;
            set { if (_selectedExportFormat == value) return; _selectedExportFormat = value; OnPropertyChanged(nameof(SelectedExportFormat)); }
        }

        private DateTime? _fromDate;
        public DateTime? FromDate
        {
            get => _fromDate;
            set { if (_fromDate == value) return; _fromDate = value; OnPropertyChanged(nameof(FromDate)); }
        }

        private DateTime? _toDate;
        public DateTime? ToDate
        {
            get => _toDate;
            set { if (_toDate == value) return; _toDate = value; OnPropertyChanged(nameof(ToDate)); }
        }

        private int _fromHour;
        public int FromHour
        {
            get => _fromHour;
            set { if (_fromHour == value) return; _fromHour = value; OnPropertyChanged(nameof(FromHour)); }
        }

        private int _fromMinute;
        public int FromMinute
        {
            get => _fromMinute;
            set { if (_fromMinute == value) return; _fromMinute = value; OnPropertyChanged(nameof(FromMinute)); }
        }

        private int _toHour = 23;
        public int ToHour
        {
            get => _toHour;
            set { if (_toHour == value) return; _toHour = value; OnPropertyChanged(nameof(ToHour)); }
        }

        private int _toMinute = 59;
        public int ToMinute
        {
            get => _toMinute;
            set { if (_toMinute == value) return; _toMinute = value; OnPropertyChanged(nameof(ToMinute)); }
        }

        private string _selectedTable;
        public string SelectedTable
        {
            get => _selectedTable;
            set
            {
                if (_selectedTable == value) return;
                _selectedTable = value;
                OnPropertyChanged(nameof(SelectedTable));
                _ = OnSelectedTableChangedAsync(value);
            }
        }

        private DataView _results;
        public DataView Results
        {
            get => _results;
            private set { _results = value; OnPropertyChanged(nameof(Results)); }
        }

        private bool _isLoading;
        public bool IsLoading
        {
            get => _isLoading;
            private set { if (_isLoading == value) return; _isLoading = value; OnPropertyChanged(nameof(IsLoading)); }
        }

        private string _busyMessage = "Please wait...";
        public string BusyMessage
        {
            get => _busyMessage;
            private set { if (_busyMessage == value) return; _busyMessage = value; OnPropertyChanged(nameof(BusyMessage)); }
        }

        public ICommand RefreshTablesCommand { get; }
        public ICommand LoadDataCommand { get; }
        public ICommand LoadDataBetweenDatesCommand { get; }
        public ICommand SelectAllColumnsCommand { get; }
        public ICommand ClearColumnsCommand { get; }
        public ICommand ExportCommand { get; }

        // Resolved logo path (automatically sourced from application folder/resources/embedded)
        public string LogoPath { get; }

        public MainViewModel(IDatabaseService db)
        {
            _db = db ?? throw new ArgumentNullException(nameof(db));
            RefreshTablesCommand = new RelayCommand(async _ => await LoadTablesAsync());
            LoadDataCommand = new RelayCommand(async _ => await LoadSelectedColumnsDataAsync());
            LoadDataBetweenDatesCommand = new RelayCommand(async _ => await LoadSelectedColumnsDataBetweenDatesAsync());
            SelectAllColumnsCommand = new RelayCommand(_ => SelectAllColumns());
            ClearColumnsCommand = new RelayCommand(_ => ClearAllColumns());
            ExportCommand = new RelayCommand(async _ => await ExportResultsAsync());

            // Resolve logo once at startup from exe folder / resources / embedded resource.
            LogoPath = ResolveLogoToTempFile();

            // sensible defaults
            ToDate = DateTime.Today.AddDays(1).Date; // include today until end
            FromDate = DateTime.Today.AddDays(-7).Date;
            FromHour = 0;
            FromMinute = 0;
            ToHour = 23;
            ToMinute = 59;
        }

        // helper to show loader and block UI while the passed async action runs
        private async Task RunWithLoading(string message, Func<Task> action)
        {
            if (IsLoading) return;
            try
            {
                BusyMessage = message;
                IsLoading = true;
                await action();
            }
            finally
            {
                IsLoading = false;
            }
        }

        public async Task InitializeAsync()
        {
            await LoadTablesAsync();
        }

        private async Task LoadTablesAsync()
        {
            string firstTable = null;

            await RunWithLoading("Loading tables...", async () =>
            {
                try
                {
                    var tables = await _db.GetTablesAsync();
                    Tables.Clear();
                    foreach (var t in tables)
                    {
                        string tableName = t.StartsWith("dbo.", StringComparison.OrdinalIgnoreCase)
                            ? t.Substring(4) : t;
                        if (firstTable == null) firstTable = tableName;
                        Tables.Add(tableName);
                    }
                    Columns.Clear();
                    Results = null;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error loading tables: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            });

            // After loading (and after IsLoading is reset), select the first table
            // so column loading can use RunWithLoading without being blocked.
            if (!string.IsNullOrWhiteSpace(firstTable))
            {
                SelectedTable = firstTable;
            }
        }

        private async Task OnSelectedTableChangedAsync(string fullName)
        {
            await RunWithLoading("Loading columns...", async () =>
            {
                Columns.Clear();
                Results = null;
                if (string.IsNullOrWhiteSpace(fullName)) return;

                var parts = fullName.Split(new[] { '.' }, 2);
                var schema = parts.Length == 2 ? parts[0] : "dbo";
                var table = parts.Length == 2 ? parts[1] : parts[0];

                try
                {
                    var cols = await _db.GetColumnsAsync(schema, table);
                    foreach (var c in cols)
                    {
                        Columns.Add(new ColumnItem { Name = c.Name, DataType = c.DataType, IsSelected = true });
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error loading columns: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            });
        }

        private async Task LoadSelectedColumnsDataAsync()
        {
            await RunWithLoading("Loading data...", async () =>
            {
                if (string.IsNullOrWhiteSpace(SelectedTable)) return;

                var selected = Columns
                    .Where(c => c.IsSelected)
                    .Select(c => c.Name)
                    .ToList();

                if (!selected.Any())
                {
                    MessageBox.Show("No columns selected.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var parts = SelectedTable.Split(new[] { '.' }, 2);
                var schema = parts.Length == 2 ? parts[0] : "dbo";
                var table = parts.Length == 2 ? parts[1] : parts[0];

                try
                {
                    var dt = await _db.GetColumnsDataAsync(schema, table, selected, 100000000);
                    Results = dt?.DefaultView;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error loading data: {ex.Message}\n\n{ex}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            });
        }

        private async Task LoadSelectedColumnsDataBetweenDatesAsync()
        {
            // validations first (don't show loader for validation checks)
            if (string.IsNullOrWhiteSpace(SelectedTable)) return;
            if (!FromDate.HasValue || !ToDate.HasValue)
            {
                MessageBox.Show("Please provide both From and To dates.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var selected = Columns
                .Where(c => c.IsSelected)
                .Select(c => c.Name)
                .ToList();

            if (!selected.Any())
            {
                MessageBox.Show("No columns selected.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var parts = SelectedTable.Split(new[] { '.' }, 2);
            var schema = parts.Length == 2 ? parts[0] : "dbo";
            var table = parts.Length == 2 ? parts[1] : parts[0];

            // perform fetch inside RunWithLoading so UI overlay is shown
            await RunWithLoading("Loading data between dates...", async () =>
            {
                try
                {
                    // Use exclusive upper bound: [TimeStamp] >= from AND [TimeStamp] < toExclusive
                    // Since UI captures up to minute precision, interpret "To" as inclusive of that minute
                    // by adding one minute to make the upper bound exclusive.
                    var from = FromDate.Value.Date
                               .AddHours(FromHour)
                               .AddMinutes(FromMinute);

                    var toInclusiveMinute = ToDate.Value.Date
                                            .AddHours(ToHour)
                                            .AddMinutes(ToMinute);

                    var toExclusive = toInclusiveMinute.AddMinutes(1);

                    // fetch all matching rows (no TOP)
                    var dt = await _db.GetColumnsDataBetweenDatesAsync(schema, table, selected, from, toExclusive);
                    Results = dt?.DefaultView;
                }
                catch (Exception ex)
                {
                    // show full exception to help diagnose (includes stack trace)
                    MessageBox.Show($"Error loading data between dates: {ex.Message}\n\n{ex}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            });
        }

        private async Task ExportResultsAsync()
        {
            // export is not considered a data-fetch operation; keep UI responsive for file dialogs.
            if (Results == null || Results.Table == null || Results.Table.Columns.Count == 0 || Results.Table.Rows.Count == 0)
            {
                MessageBox.Show("No data to export.", "Export", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var message = SelectedExportFormat == "Excel" ? "Exporting report to Excel..." : SelectedExportFormat == "PDF" ? "Exporting report to PDF..." : "Exporting report to CSV...";

            await RunWithLoading(message, async () =>
            {
                if (SelectedExportFormat == "Excel")
                {
                    await ExportHugeExcelDataAsync();
                }
                else if (SelectedExportFormat == "PDF")
                {
                    await ExportResultsToPDFAsync();
                }
                else
                {
                    await ExportResultsToCsvAsync();
                }
            });
        }

        private string ResolveLogoToTempFile()
        {
            // Read configuration first: supports "LogoCandidates" (semicolon-separated)
            // or a single "LogoPath". Relative paths are resolved against the app base dir.
            string configured = ConfigurationManager.AppSettings["LogoPath"];

            var candidates = new System.Collections.Generic.List<string>();
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;

            if (!string.IsNullOrWhiteSpace(configured))
            {
                var parts = configured.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                                      .Select(p => p.Trim());
                foreach (var p in parts)
                {
                    if (string.IsNullOrWhiteSpace(p)) continue;
                    try
                    {
                        candidates.Add(Path.IsPathRooted(p) ? p : Path.Combine(baseDir, p));
                    }
                    catch { /* skip invalid entries */ }
                }
            }
            else
            {
                // original default candidate files (relative to exe)
                candidates.Add(Path.Combine(baseDir, "logo.png"));
                candidates.Add(Path.Combine(baseDir, "Resources", "FactoryLogo.png"));
                candidates.Add(Path.Combine(baseDir, "Resources", "FactoryLogo.PNG"));
                candidates.Add(Path.Combine(baseDir, "Resources", "factorylogo.png"));
            }

            // Check file-system candidates first
            foreach (var c in candidates)
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(c) && File.Exists(c)) return c;
                }
                catch { }
            }

            // 2) try pack URI (WPF Resource)
            try
            {
                var packUri = new Uri("pack://application:,,,/Resources/FactoryLogo.png", UriKind.Absolute);
                var info = Application.GetResourceStream(packUri);
                if (info?.Stream != null)
                {
                    var temp = Path.Combine(Path.GetTempPath(), $"FactoryLogo_{Guid.NewGuid()}.png");
                    using (var fs = File.Create(temp))
                    {
                        info.Stream.CopyTo(fs);
                    }
                    return temp;
                }
            }
            catch
            {
                // ignore and continue to embedded resource attempt
            }

            // 3) try assembly manifest resource (if build action EmbeddedResource)
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                var name = asm.GetManifestResourceNames().FirstOrDefault(n => n.EndsWith("FactoryLogo.png", StringComparison.OrdinalIgnoreCase));
                if (name != null)
                {
                    using (var s = asm.GetManifestResourceStream(name))
                    {
                        if (s != null)
                        {
                            var temp = Path.Combine(Path.GetTempPath(), $"FactoryLogo_{Guid.NewGuid()}.png");
                            using (var fs = File.Create(temp)) s.CopyTo(fs);
                            return temp;
                        }
                    }
                }
            }
            catch { }

            return null;
        }

        private async Task ExportResultsToPDFAsync()
        {
            if (Results == null || Results.Table == null ||
                Results.Table.Columns.Count == 0 ||
                Results.Table.Rows.Count == 0)
            {
                MessageBox.Show("No data to export.", "Export",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*",
                FileName = $"{(SelectedTable ?? "report").Replace('.', '_')}_{DateTime.Now:yyyyMMddHHmmss}.pdf",
                DefaultExt = ".pdf"
            };

            if (dlg.ShowDialog() != true)
                return;

            var table = Results.Table.Copy(); // Important for threading safety
            var filePath = dlg.FileName;

            string tempLogoPath = null;

            try
            {
                tempLogoPath = LogoPath; // already resolved

                await Task.Run(() =>
                {
                    QuestPDF.Settings.License = LicenseType.Community;

                    Document.Create(container =>
                    {
                        container.Page(page =>
                        {
                            page.Size(PageSizes.A4.Landscape());
                            page.Margin(25);

                            // ================= HEADER =================
                            page.Header().Column(header =>
                            {
                                if (!string.IsNullOrWhiteSpace(tempLogoPath) && File.Exists(tempLogoPath))
                                {
                                    header.Item()
                                          .AlignCenter()
                                          .Height(60)
                                          .Image(tempLogoPath)
                                          .FitHeight();
                                }

                                header.Item()
                                      .AlignCenter()
                                      .Text(SelectedTable ?? "Report")
                                      .FontSize(18)
                                      .Bold();

                                string startDate = FromDate.HasValue
                                    ? FromDate.Value.Date.AddHours(FromHour).AddMinutes(FromMinute).ToString("dd-MMM-yyyy HH:mm")
                                    : "N/A";
                                string endDate = ToDate.HasValue
                                    ? ToDate.Value.Date.AddHours(ToHour).AddMinutes(ToMinute).ToString("dd-MMM-yyyy HH:mm")
                                    : "N/A";

                                header.Item()
                                      .AlignCenter()
                                      .Text($"Start Date: {startDate}   |   End Date: {endDate}")
                                      .FontSize(10);

                                header.Item().PaddingBottom(10);
                            });

                            // ================= TABLE CONTENT =================
                            page.Content().Table(tableLayout =>
                            {
                                // Define columns
                                tableLayout.ColumnsDefinition(columns =>
                                {
                                    for (int i = 0; i < table.Columns.Count; i++)
                                        columns.RelativeColumn();
                                });

                                // ===== REPEATING HEADER =====
                                tableLayout.Header(header =>
                                {
                                    foreach (DataColumn column in table.Columns)
                                    {
                                        header.Cell()
                                              .Background(Colors.Grey.Lighten2)
                                              .Border(1)
                                              .Padding(5)
                                              .Text(column.ColumnName)
                                              .Bold()
                                              .FontSize(10);
                                    }
                                });

                                // ===== DATA ROWS =====
                                foreach (DataRow row in table.Rows)
                                {
                                    foreach (var cell in row.ItemArray)
                                    {
                                        tableLayout.Cell()
                                            .Border(1)
                                            .Padding(5)
                                            .Text(cell?.ToString() ?? "")
                                            .FontSize(9);
                                    }
                                }
                            });

                            // ================= FOOTER =================
                            page.Footer().Row(footer =>
                            {
                                footer.RelativeItem()
                                      .AlignLeft()
                                      .Text($"Generated on {DateTime.Now:dd-MMM-yyyy HH:mm:ss}")
                                      .FontSize(8);

                                footer.RelativeItem()
                                      .AlignRight()
                                      .Text(x =>
                                      {
                                          x.Span("Page ").FontSize(8);
                                          x.CurrentPageNumber().FontSize(8);
                                          x.Span(" of ").FontSize(8);
                                          x.TotalPages().FontSize(8);
                                      });
                            });
                        });
                    })
                    .GeneratePdf(filePath);
                });

                MessageBox.Show(
                    $"Export complete. {table.Rows.Count} records exported to:{Environment.NewLine}{filePath}",
                    "Export",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed: {ex.Message}",
                    "Export",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                try
                {
                    // only delete temp logo files we created (those in TempPath)
                    if (!string.IsNullOrWhiteSpace(LogoPath) &&
                        LogoPath.StartsWith(Path.GetTempPath()) &&
                        File.Exists(LogoPath))
                    {
                        File.Delete(LogoPath);
                    }
                }
                catch { }
            }
        }

        private async Task ExportResultsToCsvAsync()
        {
            try
            {
                if (Results?.Table == null || Results.Table.Rows.Count == 0)
                {
                    MessageBox.Show("No data to export.");
                    return;
                }

                var dlg = new SaveFileDialog
                {
                    Filter = "CSV files (*.csv)|*.csv",
                    FileName = $"{(SelectedTable ?? "data").Replace('.', '_')}_{DateTime.Now:yyyyMMddHHmmss}.csv",
                };

                if (dlg.ShowDialog() != true)
                    return;

                var table = Results.Table.Copy(); // IMPORTANT
                var filePath = dlg.FileName;

                await Task.Run(() =>
                {
                    using (var sw = new StreamWriter(filePath, false, new UTF8Encoding(true)))
                    {
                        // header
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            if (i > 0) sw.Write(",");
                            sw.Write(table.Columns[i].ColumnName);
                        }
                        sw.WriteLine();

                        // rows
                        foreach (DataRow row in table.Rows)
                        {
                            for (int i = 0; i < table.Columns.Count; i++)
                            {
                                if (i > 0) sw.Write(",");
                                sw.Write(row[i]?.ToString());
                            }
                            sw.WriteLine();
                        }
                    }
                });

                if (!File.Exists(filePath))
                {
                    MessageBox.Show("Export failed: file was not created.", "Export", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var exportedCount = table.Rows.Count;
                MessageBox.Show($"Export complete. {exportedCount} records exported to:{Environment.NewLine}{filePath}", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private async Task ExportHugeExcelDataAsync()
        {
            try
            {
                if (Results?.Table == null || Results.Table.Rows.Count == 0)
                {
                    MessageBox.Show("No data to export.");
                    return;
                }

                var dlg = new SaveFileDialog
                {
                    Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                    FileName = $"{(SelectedTable ?? "data").Replace('.', '_')}_{DateTime.Now:yyyyMMddHHmmss}.xlsx",
                };

                if (dlg.ShowDialog() != true)
                    return;

                string filePath = dlg.FileName;
                var table = Results.Table.Copy(); // thread-safe snapshot

                await Task.Run(() =>
                {
                    const int MaxRowsPerSheet = 1048575; // Excel limit minus header

                    using (SpreadsheetDocument document =
                        SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                    {
                        WorkbookPart workbookPart = document.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();
                        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                        int totalRows = table.Rows.Count;
                        int sheetCount = (int)Math.Ceiling(totalRows / (double)MaxRowsPerSheet);

                        for (int s = 0; s < sheetCount; s++)
                        {
                            WorksheetPart worksheetPart =
                                workbookPart.AddNewPart<WorksheetPart>();

                            using (OpenXmlWriter writer = OpenXmlWriter.Create(worksheetPart))
                            {
                                writer.WriteStartElement(new Worksheet());
                                writer.WriteStartElement(new SheetData());

                                uint currentExcelRow = 1;

                                // ===== WRITE HEADER =====
                                writer.WriteStartElement(new Row { RowIndex = currentExcelRow });

                                foreach (DataColumn col in table.Columns)
                                {
                                    writer.WriteElement(new Cell
                                    {
                                        DataType = CellValues.InlineString,
                                        InlineString = new InlineString(
                                            new Text(col.ColumnName ?? ""))
                                    });
                                }

                                writer.WriteEndElement(); // Header row
                                currentExcelRow++;

                                // ===== WRITE DATA =====
                                int startRow = s * MaxRowsPerSheet;
                                int endRow = Math.Min(startRow + MaxRowsPerSheet, totalRows);

                                for (int i = startRow; i < endRow; i++)
                                {
                                    writer.WriteStartElement(new Row { RowIndex = currentExcelRow });

                                    foreach (var item in table.Rows[i].ItemArray)
                                    {
                                        writer.WriteElement(new Cell
                                        {
                                            DataType = CellValues.InlineString,
                                            InlineString = new InlineString(
                                                new Text(item?.ToString() ?? ""))
                                        });
                                    }

                                    writer.WriteEndElement(); // Data row
                                    currentExcelRow++;
                                }

                                writer.WriteEndElement(); // SheetData
                                writer.WriteEndElement(); // Worksheet
                            }

                            sheets.Append(new Sheet
                            {
                                Id = workbookPart.GetIdOfPart(worksheetPart),
                                SheetId = (uint)(s + 1),
                                Name = $"Sheet{s + 1}"
                            });
                        }

                        workbookPart.Workbook.Save();
                    }
                });

                if (!File.Exists(filePath))
                {
                    MessageBox.Show("Export failed: file was not created.", "Export", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var exportedCount = table.Rows.Count;
                MessageBox.Show($"Export complete. {exportedCount} records exported to:{Environment.NewLine}{filePath}", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
                Process.Start(new ProcessStartInfo(filePath)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed:\n{ex.Message}");
            }
        }

        private static string Escape(object value)
        {
            if (value == null || value == DBNull.Value) return string.Empty;
            var s = value.ToString();
            if (s.Contains("\"")) s = s.Replace("\"", "\"\"");
            if (s.Contains(",") || s.Contains("\r") || s.Contains("\n") || s.Contains("\""))
            {
                s = $"\"{s}\"";
            }
            return s;
        }

        private void SelectAllColumns()
        {
            foreach (var c in Columns) c.IsSelected = true;
        }

        private void ClearAllColumns()
        {
            foreach (var c in Columns) c.IsSelected = false;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        public class ColumnItem : INotifyPropertyChanged
        {
            private bool _isSelected;
            public string Name { get; set; }
            public string DataType { get; set; }
            public bool IsSelected
            {
                get => _isSelected;
                set { if (_isSelected == value) return; _isSelected = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected))); }
            }
            public event PropertyChangedEventHandler PropertyChanged;
        }
    }
}