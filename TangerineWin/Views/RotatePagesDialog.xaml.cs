using System.Windows;

namespace TangerineWin.Views;

public partial class RotatePagesDialog : Window
{
    public string PageRange =>
        RangeTextBox.Text.Trim();

    public int RotationDegrees
    {
        get
        {
            if (Rotate180Radio.IsChecked == true)
                return 180;

            if (Rotate270Radio.IsChecked == true)
                return 270;

            return 90;
        }
    }

    public RotatePagesDialog()
    {
        InitializeComponent();

        Loaded += (_, _) =>
        {
            RangeTextBox.Focus();
        };
    }

    private void OnRotateClick(
        object sender,
        RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(
            RangeTextBox.Text))
        {
            System.Windows.MessageBox.Show(
                "Введите номера страниц.",
                "Morph",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);

            return;
        }

        DialogResult = true;
    }

    private void OnCancelClick(
        object sender,
        RoutedEventArgs e)
    {
        DialogResult = false;
    }
}