using ReportGenerator.Helpers;
using ReportGenerator.Services;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using ClosedXML.Excel;

namespace ReportGenerator.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly IDatabaseService _db;

        public ObservableCollection<string> Tables { get; } = new ObservableCollection<string>();
        public ObservableCollection<ColumnItem> Columns { get; } = new ObservableCollection<ColumnItem>();
        public ObservableCollection<string> ExportFormats { get; } = new ObservableCollection<string> { "Excel", "CSV" };

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

        public MainViewModel(IDatabaseService db)
        {
            _db = db ?? throw new ArgumentNullException(nameof(db));
            RefreshTablesCommand = new RelayCommand(async _ => await LoadTablesAsync());
            LoadDataCommand = new RelayCommand(async _ => await LoadSelectedColumnsDataAsync());
            LoadDataBetweenDatesCommand = new RelayCommand(async _ => await LoadSelectedColumnsDataBetweenDatesAsync());
            SelectAllColumnsCommand = new RelayCommand(_ => SelectAllColumns());
            ClearColumnsCommand = new RelayCommand(_ => ClearAllColumns());
            ExportCommand = new RelayCommand(async _ => await ExportResultsAsync());

            // sensible defaults
            ToDate = DateTime.Today.AddDays(1).Date; // include today until end
            FromDate = DateTime.Today.AddDays(-7).Date;
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
                        if (firstTable == null) firstTable = t;
                        Tables.Add(t);
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
                    var dt = await _db.GetColumnsDataAsync(schema, table, selected, 500);
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
                    // Use exclusive upper bound: pass to = endDate.Date.AddDays(1)
                    var from = FromDate.Value.Date;
                    var toExclusive = ToDate.Value.Date.AddDays(1);

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

            var message = SelectedExportFormat == "Excel" ? "Exporting report to Excel..." : "Exporting report to CSV...";

            await RunWithLoading(message, async () =>
            {
                if (SelectedExportFormat == "Excel")
                {
                    await ExportResultsToExcelAsync();
                }
                else if(SelectedExportFormat == "PDF")
                {
                   // await ExportResultsToPDFAsync();
                }
                else
                {
                    await ExportResultsToCsvAsync();
                }
            });
        }


        private async Task ExportResultsToCsvAsync()
        {
            if (Results == null || Results.Table == null || Results.Table.Columns.Count == 0 || Results.Table.Rows.Count == 0)
            {
                MessageBox.Show("No data to export.", "Export", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                FileName = $"{(SelectedTable ?? "data").Replace('.', '_')}_{DateTime.Now:yyyyMMddHHmmss}.csv",
                DefaultExt = ".csv"
            };

            if (dlg.ShowDialog() != true) return;

            var table = Results.Table;
            var filePath = dlg.FileName;

            try
            {
                await RunWithLoading("Exporting report to CSV...", async () =>
                {
                    await Task.Run(() =>
                    {
                        using (var sw = new StreamWriter(filePath, false, new UTF8Encoding(true)))
                        {
                            // header
                            for (int i = 0; i < table.Columns.Count; i++)
                            {
                                if (i > 0) sw.Write(",");
                                sw.Write(Escape(table.Columns[i].ColumnName));
                            }
                            sw.WriteLine();

                            // rows
                            foreach (DataRow row in table.Rows)
                            {
                                for (int i = 0; i < table.Columns.Count; i++)
                                {
                                    if (i > 0) sw.Write(",");
                                    sw.Write(Escape(row[i]));
                                }
                                sw.WriteLine();
                            }
                        }
                    });
                });

                var exportedCount = table.Rows.Count;
                MessageBox.Show($"Export complete. {exportedCount} records exported to:{Environment.NewLine}{filePath}", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed: {ex.Message}", "Export", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task ExportResultsToExcelAsync()
        {
            if (Results == null || Results.Table == null || Results.Table.Columns.Count == 0 || Results.Table.Rows.Count == 0)
            {
                MessageBox.Show("No data to export.", "Export", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                FileName = $"{(SelectedTable ?? "data").Replace('.', '_')}_{DateTime.Now:yyyyMMddHHmmss}.xlsx",
                DefaultExt = ".xlsx"
            };

            if (dlg.ShowDialog() != true) return;

            var table = Results.Table;
            var filePath = dlg.FileName;

            try
            {
                await RunWithLoading("Exporting report to Excel...", async () =>
                {
                    await Task.Run(() =>
                    {
                        const int ExcelMaxRows = 1_048_576;
                        int maxDataRowsPerSheet = ExcelMaxRows - 1; // reserve one row for header
                        int totalRows = table.Rows.Count;
                        int totalSheets = (int)Math.Ceiling(totalRows / (double)maxDataRowsPerSheet);

                        using (var wb = new XLWorkbook())
                        {
                            var baseSheetName = "Data";
                            for (int s = 0; s < totalSheets; s++)
                            {
                                int start = s * maxDataRowsPerSheet;
                                int count = Math.Min(maxDataRowsPerSheet, totalRows - start);

                                // clone schema and import the chunk of rows
                                var chunk = table.Clone();
                                for (int i = 0; i < count; i++)
                                {
                                    chunk.ImportRow(table.Rows[start + i]);
                                }

                                // build safe sheet name (no invalid chars, <=31 chars)
                                string sheetName = totalSheets == 1 ? baseSheetName : $"{baseSheetName}_{s + 1}";
                                var invalid = new[] { ':', '\\', '/', '?', '*', '[', ']' };
                                foreach (var c in invalid) sheetName = sheetName.Replace(c, '_');
                                if (sheetName.Length > 31) sheetName = sheetName.Substring(0, 31);

                                var ws = wb.Worksheets.Add(chunk, sheetName);
                                ws.ColumnsUsed().AdjustToContents();
                            }

                            wb.SaveAs(filePath);
                        }
                    });
                });

                var exportedCount = table.Rows.Count;
                MessageBox.Show($"Export complete. {exportedCount} records exported to:{Environment.NewLine}{filePath}", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed: {ex.Message}", "Export", MessageBoxButton.OK, MessageBoxImage.Error);
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