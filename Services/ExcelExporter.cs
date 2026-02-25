using System.Data;
using System.Globalization;
using System.IO;
using LeakTrendViewer.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;

namespace LeakTrendViewer.Services;

public class  ExcelExporter
{
    public Task ExportAsync(string filePath, IReadOnlyList<LeakRecord> records, IReadOnlyList<string> visibleColumns)
    {
        return Task.Run(() => Export(filePath, records, visibleColumns));
    }

    public void Export(string filePath, IReadOnlyList<LeakRecord> records, IReadOnlyList<string> visibleColumns)
    {
        using var package = new ExcelPackage();
        var dataTable = BuildDataTable(records, visibleColumns);

        var ws = package.Workbook.Worksheets.Add("데이터");
        ws.Cells["A1"].LoadFromDataTable(dataTable, true);

        var endRow = Math.Max(1, dataTable.Rows.Count + 1);
        var colCount = dataTable.Columns.Count;

        ws.Cells[1, 1, 1, colCount].Style.Font.Bold = true;

        if (dataTable.Rows.Count > 0)
        {
            for (var colIndex = 1; colIndex <= colCount; colIndex++)
            {
                if (dataTable.Columns[colIndex - 1].DataType == typeof(DateTime))
                {
                    ws.Cells[2, colIndex, endRow, colIndex].Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss";
                }
            }
        }

        for (var colIndex = 1; colIndex <= colCount; colIndex++)
        {
            var column = dataTable.Columns[colIndex - 1];
            ws.Column(colIndex).Width = GetColumnWidth(column.ColumnName, column.DataType);
        }

        var range = ws.Cells[1, 1, endRow, colCount];
        range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
        range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
        range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
        range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
        ws.Cells[1, 1, 1, colCount].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        ws.Cells[endRow, 1, endRow, colCount].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        ws.Cells[1, 1, endRow, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
        ws.Cells[1, colCount, endRow, colCount].Style.Border.Right.Style = ExcelBorderStyle.Thin;

        var hasTimeColumn = visibleColumns.Contains("Data Time", StringComparer.OrdinalIgnoreCase);
        var hasLeakColumn = visibleColumns.Contains("LeakRate2", StringComparer.OrdinalIgnoreCase);
        var hasPositiveLeak = records.Any(x => x.LeakRate2 > 0d);
        if (hasTimeColumn && hasLeakColumn && hasPositiveLeak && dataTable.Rows.Count > 0)
        {
            var timeCol = dataTable.Columns["Data Time"]?.Ordinal + 1 ?? -1;
            var leakCol = dataTable.Columns["LeakRate2"]?.Ordinal + 1 ?? -1;
            if (timeCol > 0 && leakCol > 0)
            {
                var chartWs = package.Workbook.Worksheets.Add("차트");
                var chart = chartWs.Drawings.AddChart("LeakRate2Chart", eChartType.XYScatterLinesNoMarkers) as ExcelScatterChart;
                if (chart != null)
                {
                    var series = chart.Series.Add(ws.Cells[2, leakCol, endRow, leakCol], ws.Cells[2, timeCol, endRow, timeCol]);
                    series.Header = "LeakRate2";
                    chart.YAxis.LogBase = 10;
                    chart.YAxis.MinValue = 1e-12;
                    chart.Title.Text = "LeakRate2 History Trend";
                    chart.SetSize(1800, 900);
                    chart.SetPosition(1, 0, 0, 0);
                }
            }
        }

        package.SaveAs(new FileInfo(filePath));
    }

    private static DataTable BuildDataTable(IReadOnlyList<LeakRecord> records, IReadOnlyList<string> visibleColumns)
    {
        var dt = new DataTable("LeakData");
        dt.Columns.Add("NO", typeof(int));

        foreach (var columnName in visibleColumns)
        {
            dt.Columns.Add(columnName, InferColumnType(records, columnName));
        }

        dt.Columns.Add("Source", typeof(string));

        foreach (var record in records)
        {
            var row = dt.NewRow();
            row["NO"] = record.RowNo;

            foreach (var columnName in visibleColumns)
            {
                var value = record.Values.GetValueOrDefault(columnName);
                var columnType = dt.Columns[columnName]?.DataType ?? typeof(string);
                row[columnName] = NormalizeValue(value, columnType);
            }

            row["Source"] = record.SourceFile;
            dt.Rows.Add(row);
        }

        return dt;
    }

    private static object NormalizeValue(object? value, Type targetType)
    {
        if (value is null)
        {
            return DBNull.Value;
        }

        if (targetType == typeof(DateTime) && value is DateTime dateTime)
        {
            return dateTime;
        }

        if (IsNumericType(targetType) && value is IConvertible)
        {
            return Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture) ?? DBNull.Value;
        }

        return value;
    }

    private static Type InferColumnType(IReadOnlyList<LeakRecord> records, string columnName)
    {
        foreach (var record in records)
        {
            var value = record.Values.GetValueOrDefault(columnName);
            if (value is null || value is DBNull)
            {
                continue;
            }

            var valueType = value.GetType();
            if (valueType == typeof(DateTime))
            {
                return typeof(DateTime);
            }

            if (IsNumericType(valueType))
            {
                return typeof(double);
            }

            return typeof(string);
        }

        return typeof(string);
    }

    private static double GetColumnWidth(string columnName, Type columnType)
    {
        if (string.Equals(columnName, "NO", StringComparison.OrdinalIgnoreCase))
        {
            return 8;
        }

        if (string.Equals(columnName, "Source", StringComparison.OrdinalIgnoreCase))
        {
            return 28;
        }

        if (columnType == typeof(DateTime))
        {
            return 22;
        }

        return IsNumericType(columnType) ? 15 : 18;
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
}
