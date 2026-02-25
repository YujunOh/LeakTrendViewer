namespace LeakTrendViewer.Models;

public class  LeakRecord
{
    public int RowNo { get; set; }

    public DateTime Timestamp { get; set; }

    public double LeakRate2 { get; set; }

    public string SourceFile { get; set; } = string.Empty;

    /// <summary>
    /// All column values keyed by MDB column name. Used for dynamic grid display and export.
    /// </summary>
    public Dictionary<string, object?> Values { get; } = new();

    public string TimestampDisplay => Timestamp.ToString("yyyy-MM-dd HH:mm:ss");

    public string LeakRateDisplay => LeakRate2.ToString("0.###");
}
