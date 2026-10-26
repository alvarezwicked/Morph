using System.Windows;

namespace TangerineWin.Views;

public enum PdfPowerPointExportMode
{
    Editable,
    Exact
}

public partial class PdfToPowerPointDialog : Window
{
    public PdfPowerPointExportMode ExportMode
    {
        get;
        private set;
    }

    public PdfToPowerPointDialog()
    {
        InitializeComponent();

        ExportMode =
            PdfPowerPointExportMode.Editable;
    }

    private void OnExportClick(
        object sender,
        RoutedEventArgs e)
    {
        ExportMode =
            ExactRadio.IsChecked == true
                ? PdfPowerPointExportMode.Exact
                : PdfPowerPointExportMode.Editable;

        DialogResult = true;
    }
}