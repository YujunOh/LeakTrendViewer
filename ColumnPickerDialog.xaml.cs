using System.Windows;
using LeakTrendViewer.Models;

namespace LeakTrendViewer;

public partial class ColumnPickerDialog : Window
{
    public List<ColumnInfo> Columns { get; }

    public ColumnPickerDialog(List<ColumnInfo> columns)
    {
        InitializeComponent();

        Columns = columns
            .Select(x => new ColumnInfo
            {
                Name = x.Name,
                DataType = x.DataType,
                IsSelected = x.IsSelected
            })
            .ToList();

        DataContext = this;
    }

    private void SelectAllButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var column in Columns)
        {
            column.IsSelected = true;
        }

        ColumnListBox.Items.Refresh();
    }

    private void DeselectAllButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var column in Columns)
        {
            column.IsSelected = false;
        }

        ColumnListBox.Items.Refresh();
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
