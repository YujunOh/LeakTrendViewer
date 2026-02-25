using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using LeakTrendViewer.Models;
using LeakTrendViewer.Services;
using OxyPlot;
using OxyPlot.Annotations;
using OxyPlot.Axes;
using OxyPlot.Series;
using System.Windows.Input;

namespace LeakTrendViewer;

public partial class MainWindow : Window
{
    private const string DefaultDataFolder = @"C:\B_MilliData";

    private readonly ObservableCollection<LeakRecord> _records = new();
    private readonly MdbFileResolver _fileResolver = new();
    private readonly MdbDataLoader _dataLoader = new();
    private readonly CsvExporter _csvExporter = new();
    private readonly ExcelExporter _excelExporter = new();
    private readonly PlotModel _plotModel = new();
    private List<ColumnInfo> _availableColumns = new();
    private List<string> _visibleColumns = new();
    private LineAnnotation? _cursorAnnotation;
    private string _dataFolder = DefaultDataFolder;

    private string[] SearchRoots => new[] { _dataFolder };

    public MainWindow()
    {
        InitializeComponent();
        InitializePickers();
        InitializeGridAndChart();
        ScanFolder();
    }

    private void InitializePickers()
    {
        StartDatePicker.SelectedDate = DateTime.Today;
        EndDatePicker.SelectedDate = DateTime.Today;

        var hours = Enumerable.Range(0, 24).Select(x => x.ToString("D2", CultureInfo.InvariantCulture)).ToList();
        var minutes = Enumerable.Range(0, 60).Select(x => x.ToString("D2", CultureInfo.InvariantCulture)).ToList();

        StartHourCombo.ItemsSource = hours;
        EndHourCombo.ItemsSource = hours;
        StartMinuteCombo.ItemsSource = minutes;
        EndMinuteCombo.ItemsSource = minutes;

        StartHourCombo.SelectedIndex = 0;
        StartMinuteCombo.SelectedIndex = 0;
        EndHourCombo.SelectedIndex = DateTime.Now.Hour;
        EndMinuteCombo.SelectedIndex = DateTime.Now.Minute;
    }

    private void InitializeGridAndChart()
    {
        LeakGrid.ItemsSource = _records;
        RebuildGridColumns();

        _plotModel.Title = "History Trend";
        _plotModel.IsLegendVisible = true;
        _plotModel.PlotAreaBorderColor = OxyColors.SteelBlue;

        var xAxis = new DateTimeAxis
        {
            Position = AxisPosition.Bottom,
            StringFormat = "MM-dd HH:mm:ss",
            IntervalType = DateTimeIntervalType.Minutes,
            IsPanEnabled = true,
            IsZoomEnabled = true,
            MajorGridlineStyle = LineStyle.Solid,
            MinorGridlineStyle = LineStyle.Dot
        };

        var yAxis = new LogarithmicAxis
        {
            Position = AxisPosition.Left,
            Base = 10,
            Minimum = 1e-12,
            Maximum = 1e3,
            Title = "LeakRate2 [Pa]",
            IsPanEnabled = true,
            IsZoomEnabled = true,
            MajorGridlineStyle = LineStyle.Solid,
            MinorGridlineStyle = LineStyle.Dot
        };

        _plotModel.Axes.Add(xAxis);
        _plotModel.Axes.Add(yAxis);
        _plotModel.Series.Add(new LineSeries
        {
            Title = "LeakRate2",
            StrokeThickness = 1.3,
            Color = OxyColor.FromRgb(39, 116, 174),
            TrackerFormatString = "{2:yyyy-MM-dd HH:mm:ss}\nLeakRate2: {4:0.###}"
        });

        _cursorAnnotation = new LineAnnotation
        {
            Type = LineAnnotationType.Vertical,
            Color = OxyColor.FromRgb(255, 188, 0),
            StrokeThickness = 2,
            X = DateTimeAxis.ToDouble(DateTime.Now),
            Text = "Selected"
        };
        _plotModel.Annotations.Add(_cursorAnnotation);

        TrendPlotView.Model = _plotModel;
    }

    private async void LoadButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var start = GetDateTime(StartDatePicker, StartHourCombo, StartMinuteCombo);
            var end = GetDateTime(EndDatePicker, EndHourCombo, EndMinuteCombo);
            if (end < start)
            {
                SetStatus("시작 시간이 종료 시간보다 늦습니다.");
                return;
            }

            var resolved = _fileResolver.ResolveFiles(SearchRoots, start, end);
            if (resolved.ExistingFiles.Count == 0)
            {
                SetStatus($"파일이 없습니다. 예상 첫 파일: {resolved.AllCandidates.FirstOrDefault()}");
                _records.Clear();
                RefreshPlot();
                return;
            }

            var result = await LoadFromFilesAsync(resolved.ExistingFiles, start, end, $"{resolved.ExistingFiles.Count}개 파일 로드 중...");

            var warningText = string.Join(" | ", result.Warnings);
            var missedText = resolved.MissingFiles.Count == 0
                ? string.Empty
                : $" 누락 파일 {resolved.MissingFiles.Count}개.";

            SetStatus($"완료: {_records.Count}건{missedText}" + (string.IsNullOrWhiteSpace(warningText) ? string.Empty : $" 경고: {warningText}"));
        }
        catch (Exception ex)
        {
            SetStatus($"조회 실패: {ex.Message}");
        }
        finally
        {
            LoadButton.IsEnabled = true;
        }
    }

    private async void ImportFilesButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var start = GetDateTime(StartDatePicker, StartHourCombo, StartMinuteCombo);
            var end = GetDateTime(EndDatePicker, EndHourCombo, EndMinuteCombo);
            if (end < start)
            {
                SetStatus("시작 시간이 종료 시간보다 늦습니다.");
                return;
            }

            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Access DB (*.mdb)|*.mdb",
                Multiselect = true,
                CheckFileExists = true,
                Title = "MDB 파일 선택"
            };

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            var uniqueFiles = dialog.FileNames.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (uniqueFiles.Count == 0)
            {
                SetStatus("선택된 파일이 없습니다.");
                return;
            }

            var result = await LoadFromFilesAsync(uniqueFiles, start, end, $"선택 파일 {uniqueFiles.Count}개 로드 중...");
            var warningText = string.Join(" | ", result.Warnings);
            SetStatus($"완료: {_records.Count}건" + (string.IsNullOrWhiteSpace(warningText) ? string.Empty : $" 경고: {warningText}"));
        }
        catch (Exception ex)
        {
            SetStatus($"파일 불러오기 실패: {ex.Message}");
        }
    }

    private async Task<LoadResult> LoadFromFilesAsync(IReadOnlyList<string> files, DateTime start, DateTime end, string progressMessage)
    {
        SetStatus(progressMessage);
        LoadButton.IsEnabled = false;
        ImportFilesButton.IsEnabled = false;

        try
        {
            var result = await _dataLoader.LoadAsync(files, start, end);
            _records.Clear();

            var rowNo = 1;
            foreach (var item in result.Records)
            {
                item.RowNo = rowNo++;
                _records.Add(item);
            }

            if (_availableColumns.Count == 0)
            {
                var discoveredColumns = await _dataLoader.DiscoverColumnsAsync(files);
                _availableColumns = discoveredColumns
                    .Select(x => new ColumnInfo
                    {
                        Name = x.Name,
                        DataType = x.DataType,
                        IsSelected = string.Equals(x.Name, "Data Time", StringComparison.OrdinalIgnoreCase)
                                     || string.Equals(x.Name, "LeakRate2", StringComparison.OrdinalIgnoreCase)
                    })
                    .ToList();
            }

            _visibleColumns = _availableColumns
                .Where(x => x.IsSelected)
                .Select(x => x.Name)
                .ToList();

            RebuildGridColumns();

            RefreshPlot();
            return result;
        }
        finally
        {
            LoadButton.IsEnabled = true;
            ImportFilesButton.IsEnabled = true;
        }
    }

    private static DateTime GetDateTime(DatePicker datePicker, ComboBox hourCombo, ComboBox minuteCombo)
    {
        if (datePicker.SelectedDate is null)
        {
            throw new InvalidOperationException("날짜를 선택해주세요.");
        }

        var hour = hourCombo.SelectedIndex < 0 ? 0 : hourCombo.SelectedIndex;
        var minute = minuteCombo.SelectedIndex < 0 ? 0 : minuteCombo.SelectedIndex;

        return new DateTime(
            datePicker.SelectedDate.Value.Year,
            datePicker.SelectedDate.Value.Month,
            datePicker.SelectedDate.Value.Day,
            hour,
            minute,
            0);
    }

    private void RefreshPlot()
    {
        var lineSeries = (LineSeries)_plotModel.Series[0];
        lineSeries.Points.Clear();

        var yAxis = (LogarithmicAxis)_plotModel.Axes[1];

        var positiveValues = _records.Where(r => r.LeakRate2 > 0).Select(r => r.LeakRate2).ToList();
        double logFloor;

        if (positiveValues.Count > 0)
        {
            var dataMin = positiveValues.Min();
            var dataMax = positiveValues.Max();
            logFloor = dataMin / 10;
            yAxis.Minimum = logFloor;
            yAxis.Maximum = dataMax * 10;
        }
        else
        {
            logFloor = 1e-13;
            yAxis.Minimum = 1e-12;
            yAxis.Maximum = 1e3;
        }

        foreach (var record in _records)
        {
            var y = record.LeakRate2 > 0 ? record.LeakRate2 : logFloor;
            lineSeries.Points.Add(new DataPoint(DateTimeAxis.ToDouble(record.Timestamp), y));
        }

        if (_records.Count > 0)
        {
            var min = DateTimeAxis.ToDouble(_records.Min(x => x.Timestamp));
            var max = DateTimeAxis.ToDouble(_records.Max(x => x.Timestamp));
            _plotModel.Axes[0].Zoom(min, max);
            if (_cursorAnnotation != null)
            {
                _cursorAnnotation.X = min;
            }
        }

        _plotModel.InvalidatePlot(true);
    }

    private void LeakGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (LeakGrid.SelectedItem is not LeakRecord selected || _cursorAnnotation is null)
        {
            return;
        }

        _cursorAnnotation.X = DateTimeAxis.ToDouble(selected.Timestamp);
        _plotModel.InvalidatePlot(false);
    }

    private void TrendPlotView_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (_records.Count == 0)
        {
            return;
        }

        var position = e.GetPosition(TrendPlotView);
        var screenPoint = new ScreenPoint(position.X, position.Y);
        var xAxis = _plotModel.Axes[0];
        var yAxis = _plotModel.Axes[1];
        var dataPoint = Axis.InverseTransform(screenPoint, xAxis, yAxis);
        var clickedTime = DateTimeAxis.ToDateTime(dataPoint.X, TimeSpan.FromSeconds(1));

        var closest = _records.OrderBy(r => Math.Abs((r.Timestamp - clickedTime).TotalSeconds)).FirstOrDefault();
        if (closest == null)
        {
            return;
        }

        LeakGrid.SelectedItem = closest;
        LeakGrid.ScrollIntoView(closest);
    }

    private void SyncButton_Click(object sender, RoutedEventArgs e)
    {
        var selected = LeakGrid.SelectedItems.Cast<LeakRecord>().ToList();
        if (selected.Count == 0)
        {
            SetStatus("동기화할 선택 행이 없습니다.");
            return;
        }

        var xAxis = _plotModel.Axes[0];
        var min = DateTimeAxis.ToDouble(selected.Min(s => s.Timestamp));
        var max = DateTimeAxis.ToDouble(selected.Max(s => s.Timestamp));
        if (Math.Abs(max - min) < 0.0000001)
        {
            max += 1d / (24 * 60);
        }

        xAxis.Zoom(min, max);
        _plotModel.InvalidatePlot(false);
        SetStatus("차트 시간 범위를 선택 행 기준으로 동기화했습니다.");
    }

    private void RebuildGridColumns()
    {
        LeakGrid.Columns.Clear();

        LeakGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "NO",
            Width = 70,
            Binding = new Binding("RowNo")
        });

        foreach (var columnName in _visibleColumns)
        {
            var binding = new Binding($"Values[{columnName}]");
            var columnInfo = _availableColumns.FirstOrDefault(x => string.Equals(x.Name, columnName, StringComparison.OrdinalIgnoreCase));
            if (columnInfo?.DataType == typeof(DateTime))
            {
                binding.StringFormat = "yyyy-MM-dd HH:mm:ss";
            }
            else if (columnInfo != null && IsNumericType(columnInfo.DataType))
            {
                binding.StringFormat = "0.###";
            }

            LeakGrid.Columns.Add(new DataGridTextColumn
            {
                Header = columnName,
                Width = 160,
                Binding = binding
            });
        }

        LeakGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Source",
            Width = new DataGridLength(1, DataGridLengthUnitType.Star),
            Binding = new Binding("SourceFile")
        });
    }

    private static bool IsNumericType(Type type)
    {
        var actualType = Nullable.GetUnderlyingType(type) ?? type;
        return actualType == typeof(byte)
            || actualType == typeof(sbyte)
            || actualType == typeof(short)
            || actualType == typeof(ushort)
            || actualType == typeof(int)
            || actualType == typeof(uint)
            || actualType == typeof(long)
            || actualType == typeof(ulong)
            || actualType == typeof(float)
            || actualType == typeof(double)
            || actualType == typeof(decimal);
    }

    private void ColumnPickerButton_Click(object sender, RoutedEventArgs e)
    {
        if (_availableColumns.Count == 0)
        {
            SetStatus("먼저 데이터를 로드해주세요");
            return;
        }

        var dialog = new ColumnPickerDialog(_availableColumns)
        {
            Owner = this
        };

        if (dialog.ShowDialog() == true)
        {
            _availableColumns = dialog.Columns
                .Select(x => new ColumnInfo
                {
                    Name = x.Name,
                    DataType = x.DataType,
                    IsSelected = x.IsSelected
                })
                .ToList();

            _visibleColumns = _availableColumns
                .Where(x => x.IsSelected)
                .Select(x => x.Name)
                .ToList();

            RebuildGridColumns();
        }
    }

    private void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        if (_records.Count == 0)
        {
            SetStatus("내보낼 데이터가 없습니다.");
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "CSV file (*.csv)|*.csv",
            FileName = $"LeakRate2_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        _csvExporter.Export(dialog.FileName, _records.ToList(), _visibleColumns);
        SetStatus($"CSV 저장 완료: {Path.GetFileName(dialog.FileName)}");
    }

    private async void ExportExcelButton_Click(object sender, RoutedEventArgs e)
    {
        if (_records.Count == 0)
        {
            SetStatus("내보낼 데이터가 없습니다.");
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel file (*.xlsx)|*.xlsx",
            FileName = $"LeakRate2_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            var visibleCols = _visibleColumns.Count == 0
                ? new List<string> { "Data Time", "LeakRate2" }
                : _visibleColumns;

            await _excelExporter.ExportAsync(dialog.FileName, _records.ToList(), visibleCols);
            SetStatus($"Excel 저장 완료: {Path.GetFileName(dialog.FileName)}");

            Process.Start(new ProcessStartInfo
            {
                FileName = dialog.FileName,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            SetStatus($"Excel 저장 실패: {ex.Message}");
        }
    }

    // ── Folder browsing & scanning ──

    private void BrowseFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "MDB 파일이 있는 폴더에서 아무 파일이나 선택하세요",
            Filter = "MDB files (*.mdb)|*.mdb|All files (*.*)|*.*",
            CheckFileExists = false,
            FileName = "폴더 선택"
        };

        if (!string.IsNullOrEmpty(_dataFolder) && Directory.Exists(_dataFolder))
        {
            dialog.InitialDirectory = _dataFolder;
        }

        if (dialog.ShowDialog(this) == true)
        {
            var selectedDir = Path.GetDirectoryName(dialog.FileName);
            if (!string.IsNullOrWhiteSpace(selectedDir) && Directory.Exists(selectedDir))
            {
                _dataFolder = selectedDir;
                DataFolderTextBox.Text = _dataFolder;
                ScanFolder();
                UpdateFileCount();
                SetStatus($"데이터 폴더 변경: {_dataFolder}");
            }
        }
    }

    private void ScanFolder()
    {
        try
        {
            if (!Directory.Exists(_dataFolder))
            {
                FolderInfoLabel.Text = "(폴더 없음)";
                FolderInfoLabel.Foreground = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(211, 47, 47));
                return;
            }

            var mdbFiles = Directory.GetFiles(_dataFolder, "*.mdb");
            var count = mdbFiles.Length;

            if (count == 0)
            {
                FolderInfoLabel.Text = "(MDB 파일 없음)";
                FolderInfoLabel.Foreground = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(211, 47, 47));
                return;
            }

            // Find date range of files
            DateTime? earliest = null;
            DateTime? latest = null;
            foreach (var file in mdbFiles)
            {
                var name = Path.GetFileNameWithoutExtension(file);
                if (name.Length >= 8 && DateTime.TryParseExact(name[..8], "yyyyMMdd",
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out var fileDate))
                {
                    if (earliest is null || fileDate < earliest) earliest = fileDate;
                    if (latest is null || fileDate > latest) latest = fileDate;
                }
            }

            var rangeText = earliest.HasValue && latest.HasValue
                ? $" ({earliest:yyyy-MM-dd} ~ {latest:yyyy-MM-dd})"
                : "";

            FolderInfoLabel.Text = $"총 {count}개 MDB 파일{rangeText}";
            FolderInfoLabel.Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(46, 125, 50));
        }
        catch
        {
            FolderInfoLabel.Text = "(스캔 실패)";
        }
    }

    // ── Quick date presets ──

    private void PresetToday_Click(object sender, RoutedEventArgs e)
    {
        SetDateRange(DateTime.Today, 0, 0, DateTime.Today, 23, 59);
    }

    private void PresetYesterday_Click(object sender, RoutedEventArgs e)
    {
        var yesterday = DateTime.Today.AddDays(-1);
        SetDateRange(yesterday, 0, 0, yesterday, 23, 59);
    }

    private void PresetLast3Days_Click(object sender, RoutedEventArgs e)
    {
        SetDateRange(DateTime.Today.AddDays(-2), 0, 0, DateTime.Today, 23, 59);
    }

    private void PresetLast7Days_Click(object sender, RoutedEventArgs e)
    {
        SetDateRange(DateTime.Today.AddDays(-6), 0, 0, DateTime.Today, 23, 59);
    }

    private void PresetLast30Days_Click(object sender, RoutedEventArgs e)
    {
        SetDateRange(DateTime.Today.AddDays(-29), 0, 0, DateTime.Today, 23, 59);
    }

    private void PresetLastHour_Click(object sender, RoutedEventArgs e)
    {
        var now = DateTime.Now;
        var oneHourAgo = now.AddHours(-1);
        SetDateRange(oneHourAgo.Date, oneHourAgo.Hour, oneHourAgo.Minute, now.Date, now.Hour, now.Minute);
    }

    private void SetDateRange(DateTime startDate, int startHour, int startMinute, DateTime endDate, int endHour, int endMinute)
    {
        StartDatePicker.SelectedDate = startDate;
        EndDatePicker.SelectedDate = endDate;
        StartHourCombo.SelectedIndex = startHour;
        StartMinuteCombo.SelectedIndex = startMinute;
        EndHourCombo.SelectedIndex = endHour;
        EndMinuteCombo.SelectedIndex = endMinute;
        UpdateFileCount();
    }

    // ── File count preview ──

    private void DatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateFileCount();
    }

    private void UpdateFileCount()
    {
        try
        {
            if (StartDatePicker.SelectedDate == null || EndDatePicker.SelectedDate == null)
            {
                FileCountLabel.Text = "";
                return;
            }

            var start = StartDatePicker.SelectedDate.Value;
            var end = EndDatePicker.SelectedDate.Value;
            if (end < start)
            {
                FileCountLabel.Text = "";
                return;
            }

            var resolved = _fileResolver.ResolveFiles(SearchRoots, start, end);
            FileCountLabel.Text = $"[MDB 파일: {resolved.ExistingFiles.Count}개 / {resolved.AllCandidates.Count}일]";
        }
        catch
        {
            FileCountLabel.Text = "";
        }
    }

    // ── Excel Report (direct export) ──

    private async void ExcelReportButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var start = GetDateTime(StartDatePicker, StartHourCombo, StartMinuteCombo);
            var end = GetDateTime(EndDatePicker, EndHourCombo, EndMinuteCombo);
            if (end < start)
            {
                SetStatus("시작 시간이 종료 시간보다 늦습니다.");
                return;
            }

            var resolved = _fileResolver.ResolveFiles(SearchRoots, start, end);
            if (resolved.ExistingFiles.Count == 0)
            {
                SetStatus($"해당 날짜 범위에 MDB 파일이 없습니다.");
                return;
            }

            ExcelReportButton.IsEnabled = false;
            SetStatus($"{resolved.ExistingFiles.Count}개 파일에서 데이터 로딩 중...");

            var loadResult = await _dataLoader.LoadAsync(resolved.ExistingFiles, start, end);

            if (loadResult.Records.Count == 0)
            {
                SetStatus("해당 시간 범위에 데이터가 없습니다.");
                return;
            }

            var rowNo = 1;
            foreach (var record in loadResult.Records)
            {
                record.RowNo = rowNo++;
            }

            var outputDir = @"C:\B_MilliData";
            if (!Directory.Exists(outputDir))
            {
                outputDir = Path.GetTempPath();
            }

            var fileName = $"Report_{start:yyyyMMdd_HHmm}_{end:yyyyMMdd_HHmm}.xlsx";
            var outputPath = Path.Combine(outputDir, fileName);

            SetStatus($"{loadResult.Records.Count}건 Excel 내보내기 중...");

            var visibleCols = _visibleColumns.Count == 0
                ? new List<string> { "Data Time", "LeakRate2" }
                : _visibleColumns;

            await _excelExporter.ExportAsync(outputPath, loadResult.Records, visibleCols);

            SetStatus($"Excel Report 완료: {fileName} ({loadResult.Records.Count}건, {resolved.ExistingFiles.Count}개 파일)");

            Process.Start(new ProcessStartInfo
            {
                FileName = outputPath,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            SetStatus($"Excel Report 실패: {ex.Message}");
        }
        finally
        {
            ExcelReportButton.IsEnabled = true;
        }
    }

    private void SetStatus(string message)
    {
        StatusText.Text = message;
    }
}
