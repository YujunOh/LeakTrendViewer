using System.Text;
using System.IO;
using System.Globalization;
using LeakTrendViewer.Models;

namespace LeakTrendViewer.Services;

public class  CsvExporter
{
    public void Export(string filePath, IReadOnlyList<LeakRecord> records, IReadOnlyList<string> visibleColumns)
    {
        using var writer = new StreamWriter(filePath, false, new UTF8Encoding(true));
        var headerParts = new List<string> { "NO" };
        headerParts.AddRange(visibleColumns.Select(Escape));
        headerParts.Add("Source");
        writer.WriteLine(string.Join(",", headerParts));

        foreach (var record in records)
        {
            var rowParts = new List<string> { record.RowNo.ToString() };
            foreach (var columnName in visibleColumns)
            {
                var value = record.Values.GetValueOrDefault(columnName);
                rowParts.Add(Escape(FormatValue(value)));
            }

            rowParts.Add(Escape(record.SourceFile));
            writer.WriteLine(string.Join(",", rowParts));
        }
    }

    private static string FormatValue(object? value)
    {
        return value switch
        {
            null => string.Empty,
            DateTime dateTime => dateTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
            double number => number.ToString("0.###", CultureInfo.InvariantCulture),
            float number => number.ToString("0.###", CultureInfo.InvariantCulture),
            decimal number => number.ToString("0.###", CultureInfo.InvariantCulture),
            _ => value.ToString() ?? string.Empty
        };
    }

    private static string Escape(string value)
    {
        if (!value.Contains(',') && !value.Contains('"'))
        {
            return value;
        }

        return $"\"{value.Replace("\"", "\"\"")}\"";
    }
}
