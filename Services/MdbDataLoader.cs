using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using LeakTrendViewer.Models;

namespace LeakTrendViewer.Services;

public class  MdbDataLoader
{
    private static readonly string[] TimeNameHints = ["time", "date", "datetime", "일시", "시간"];
    private static readonly string[] LeakNameHints = ["leakrate2", "leak", "rate", "누설"];

    public async Task<LoadResult> LoadAsync(IEnumerable<string> files, DateTime start, DateTime end)
    {
        var result = new LoadResult();

        foreach (var file in files)
        {
            try
            {
                var records = await Task.Run(() => LoadFile(file, start, end));
                result.Records.AddRange(records);
            }
            catch (OleDbException ex)
            {
                result.Warnings.Add($"{Path.GetFileName(file)} OLEDB 오류: {ex.Message}");
            }
            catch (Exception ex)
            {
                result.Warnings.Add($"{Path.GetFileName(file)} 로드 실패: {ex.Message}");
            }
        }

        result.Records.Sort((a, b) => a.Timestamp.CompareTo(b.Timestamp));
        return result;
    }

    public async Task<List<ColumnInfo>> DiscoverColumnsAsync(IEnumerable<string> files)
    {
        var firstFile = files.FirstOrDefault(static x => !string.IsNullOrWhiteSpace(x));
        if (string.IsNullOrWhiteSpace(firstFile))
        {
            return [];
        }

        return await Task.Run(() => DiscoverColumnsFromFile(firstFile));
    }

    private static List<LeakRecord> LoadFile(string filePath, DateTime start, DateTime end)
    {
        var pathToUse = filePath;
        var cleanupTemp = false;

        try
        {
            using var connection = OpenConnection(pathToUse);
            return ReadRecordsFromConnection(connection, pathToUse, start, end);
        }
        catch (OleDbException)
        {
            pathToUse = CopyToTemp(filePath);
            cleanupTemp = true;
            using var fallbackConnection = OpenConnection(pathToUse);
            return ReadRecordsFromConnection(fallbackConnection, filePath, start, end);
        }
        finally
        {
            if (cleanupTemp && File.Exists(pathToUse))
            {
                File.Delete(pathToUse);
            }
        }
    }

    private static List<ColumnInfo> DiscoverColumnsFromFile(string filePath)
    {
        var pathToUse = filePath;
        var cleanupTemp = false;

        try
        {
            using var connection = OpenConnection(pathToUse);
            return DiscoverColumnsFromConnection(connection);
        }
        catch (OleDbException)
        {
            pathToUse = CopyToTemp(filePath);
            cleanupTemp = true;
            using var fallbackConnection = OpenConnection(pathToUse);
            return DiscoverColumnsFromConnection(fallbackConnection);
        }
        finally
        {
            if (cleanupTemp && File.Exists(pathToUse))
            {
                File.Delete(pathToUse);
            }
        }
    }

    private static string CopyToTemp(string sourceFilePath)
    {
        var tempPath = Path.Combine(Path.GetTempPath(), $"LeakTrend_{Guid.NewGuid():N}.mdb");
        File.Copy(sourceFilePath, tempPath, true);
        return tempPath;
    }

    private static OleDbConnection OpenConnection(string filePath)
    {
        var attempts = new[]
        {
            $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Mode=Read;Persist Security Info=False;",
            $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={filePath};Mode=Read;Persist Security Info=False;"
        };

        Exception? lastError = null;

        foreach (var connectionString in attempts)
        {
            try
            {
                var connection = new OleDbConnection(connectionString);
                connection.Open();
                return connection;
            }
            catch (Exception ex)
            {
                lastError = ex;
            }
        }

        throw new InvalidOperationException($"OLEDB 드라이버를 찾을 수 없습니다. 파일: {filePath}", lastError);
    }

    private static List<LeakRecord> ReadRecordsFromConnection(OleDbConnection connection, string sourceLabel, DateTime start, DateTime end)
    {
        foreach (var tableName in GetTableNames(connection))
        {
            if (!TryDiscoverColumns(connection, tableName, out var timeColumn, out var leakColumn))
            {
                continue;
            }

            if (!TryGetColumnMetas(connection, tableName, out var columnMetas))
            {
                continue;
            }

            var selectColumns = string.Join(", ", columnMetas.Select(x => $"[{EscapeIdentifier(x.Name)}]"));

            var sql = $"SELECT {selectColumns} FROM [{EscapeIdentifier(tableName)}] WHERE [{EscapeIdentifier(timeColumn)}] >= ? AND [{EscapeIdentifier(timeColumn)}] <= ? ORDER BY [{EscapeIdentifier(timeColumn)}] ASC";
            using var command = new OleDbCommand(sql, connection);
            command.Parameters.AddWithValue("@p1", start);
            command.Parameters.AddWithValue("@p2", end);

            using var reader = command.ExecuteReader();
            if (reader == null)
            {
                continue;
            }

            var timeOrdinal = reader.GetOrdinal(timeColumn);
            var leakOrdinal = reader.GetOrdinal(leakColumn);
            var columnOrdinals = columnMetas
                .Select(x => (x.Name, Ordinal: reader.GetOrdinal(x.Name)))
                .ToList();

            var rows = new List<LeakRecord>();
            while (reader.Read())
            {
                if (reader.IsDBNull(timeOrdinal))
                {
                    continue;
                }

                var timestamp = Convert.ToDateTime(reader.GetValue(timeOrdinal), CultureInfo.InvariantCulture);
                var leakRate = reader.IsDBNull(leakOrdinal)
                    ? 0d
                    : Convert.ToDouble(reader.GetValue(leakOrdinal), CultureInfo.InvariantCulture);

                var record = new LeakRecord
                {
                    Timestamp = timestamp,
                    LeakRate2 = leakRate,
                    SourceFile = Path.GetFileName(sourceLabel)
                };

                foreach (var (columnName, ordinal) in columnOrdinals)
                {
                    record.Values[columnName] = reader.IsDBNull(ordinal) ? null : reader.GetValue(ordinal);
                }

                rows.Add(record);
            }

            if (rows.Count > 0)
            {
                return rows;
            }
        }

        return [];
    }

    private static List<ColumnInfo> DiscoverColumnsFromConnection(OleDbConnection connection)
    {
        foreach (var tableName in GetTableNames(connection))
        {
            if (!TryGetColumnMetas(connection, tableName, out var columnMetas))
            {
                continue;
            }

            return columnMetas
                .Select(x => new ColumnInfo
                {
                    Name = x.Name,
                    DataType = x.DataType ?? typeof(object),
                    IsSelected = true
                })
                .ToList();
        }

        return [];
    }

    private static IEnumerable<string> GetTableNames(OleDbConnection connection)
    {
        using var tableSchema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        if (tableSchema is null)
        {
            yield break;
        }

        foreach (DataRow row in tableSchema.Rows)
        {
            var tableType = row["TABLE_TYPE"]?.ToString();
            var tableName = row["TABLE_NAME"]?.ToString();
            if (!string.Equals(tableType, "TABLE", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!string.IsNullOrWhiteSpace(tableName))
            {
                yield return tableName;
            }
        }
    }

    private static bool TryDiscoverColumns(OleDbConnection connection, string tableName, out string timeColumn, out string leakColumn)
    {
        timeColumn = string.Empty;
        leakColumn = string.Empty;

        if (!TryGetColumnMetas(connection, tableName, out var columnMetas))
        {
            return false;
        }

        timeColumn = columnMetas
            .Where(x => x.DataType == typeof(DateTime) || NameLooksLike(x.Name, TimeNameHints))
            .Select(x => x.Name)
            .FirstOrDefault() ?? string.Empty;

        leakColumn = columnMetas
            .Where(x => x.DataType == typeof(double) || x.DataType == typeof(float) || x.DataType == typeof(decimal) || x.DataType == typeof(int) || NameLooksLike(x.Name, LeakNameHints))
            .OrderByDescending(x => x.Name.Contains("LeakRate2", StringComparison.OrdinalIgnoreCase))
            .Select(x => x.Name)
            .FirstOrDefault() ?? string.Empty;

        return !string.IsNullOrWhiteSpace(timeColumn) && !string.IsNullOrWhiteSpace(leakColumn);
    }

    private static bool TryGetColumnMetas(OleDbConnection connection, string tableName, out List<(string Name, Type? DataType)> columnMetas)
    {
        columnMetas = new List<(string Name, Type? DataType)>();

        using var command = new OleDbCommand($"SELECT TOP 1 * FROM [{EscapeIdentifier(tableName)}]", connection);
        using var reader = command.ExecuteReader(CommandBehavior.SchemaOnly);
        if (reader is null)
        {
            return false;
        }

        var schema = reader.GetSchemaTable();
        if (schema == null)
        {
            return false;
        }

        columnMetas = schema.Rows.Cast<DataRow>()
            .Select(x => new
            {
                Name = x["ColumnName"]?.ToString() ?? string.Empty,
                DataType = x["DataType"] as Type
            })
            .Where(x => !string.IsNullOrWhiteSpace(x.Name))
            .Select(x => (x.Name, x.DataType))
            .ToList();

        return columnMetas.Count > 0;
    }

    private static bool NameLooksLike(string name, IEnumerable<string> hints)
        => hints.Any(hint => name.Contains(hint, StringComparison.OrdinalIgnoreCase));

    private static string EscapeIdentifier(string value)
        => value.Replace("]", "]]", StringComparison.Ordinal);
}

public class LoadResult
{
    public List<LeakRecord> Records { get; } = new();

    public List<string> Warnings { get; } = new();
}
