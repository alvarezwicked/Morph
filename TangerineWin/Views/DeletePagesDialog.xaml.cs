using System.Windows;

namespace TangerineWin.Views;

public partial class DeletePagesDialog : Window
{
    public string PageRange =>
        RangeTextBox.Text.Trim();

    public DeletePagesDialog()
    {
        InitializeComponent();

        Loaded += (_, _) =>
        {
            RangeTextBox.Focus();
        };
    }

    private void OnDeleteClick(
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