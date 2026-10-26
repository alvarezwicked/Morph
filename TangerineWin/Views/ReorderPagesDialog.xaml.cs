using System.Windows;

namespace TangerineWin.Views;

public partial class ReorderPagesDialog : Window
{
    public string PageOrder =>
        OrderTextBox.Text.Trim();

    public ReorderPagesDialog()
    {
        InitializeComponent();

        Loaded += (_, _) =>
        {
            OrderTextBox.Focus();
        };
    }

    private void OnReorderClick(
        object sender,
        RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(
            OrderTextBox.Text))
        {
            System.Windows.MessageBox.Show(
                "Введите порядок страниц.",
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