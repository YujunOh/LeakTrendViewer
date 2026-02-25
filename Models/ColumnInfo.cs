namespace LeakTrendViewer.Models;

public class  ColumnInfo
{
    public string Name { get; set; } = string.Empty;

    public Type DataType { get; set; } = typeof(object);

    public bool IsSelected { get; set; }
}
