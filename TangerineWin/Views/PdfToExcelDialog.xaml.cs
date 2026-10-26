using System.Windows;

namespace TangerineWin.Views;

public enum PdfExcelExportMode
{
    TablesOnly,
    TextAndTables
}

public partial class PdfToExcelDialog : Window
{
    public PdfExcelExportMode ExportMode
    {
        get;
        private set;
    }

    public PdfToExcelDialog()
    {
        InitializeComponent();

        ExportMode =
            PdfExcelExportMode.TextAndTables;
    }

    private void OnExportClick(
        object sender,
        RoutedEventArgs e)
    {
        ExportMode =
            TablesOnlyRadio.IsChecked == true
                ? PdfExcelExportMode.TablesOnly
                : PdfExcelExportMode.TextAndTables;

        DialogResult = true;
    }
}