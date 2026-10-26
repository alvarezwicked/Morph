using System.Windows;

namespace TangerineWin.Views;

public partial class PageRangeDialog : Window
{
    public string PageRange =>
        RangeTextBox.Text.Trim();

    public PageRangeDialog()
    {
        InitializeComponent();

        Loaded += (_, _) =>
        {
            RangeTextBox.Focus();
        };
    }

    private void OnExtractClick(
        object sender,
        RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(RangeTextBox.Text))
        {
            System.Windows.MessageBox.Show(
                "Введите номера страниц.",
                "Morph",
                MessageBoxButton.OK,
                MessageBoxImage.Information);

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