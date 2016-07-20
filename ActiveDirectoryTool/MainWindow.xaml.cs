using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Extensions.PrimitiveExtensions;

namespace ActiveDirectoryTool
{
    public partial class MainWindow
    {
        private void DataGrid_AutoGeneratingColumn(
            object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            e.Column.Header = e.Column.Header.ToString().SpaceCamelCase();
            if (ShowDistinguishedNamesCheckbox.IsChecked != null && 
                ShowDistinguishedNamesCheckbox.IsChecked.Value) return;
            if (e.Column.Header.ToString().Contains("Distinguished Name"))
                e.Column.Visibility = Visibility.Hidden;
        }

        private void Label_MouseDoubleClick(
            object sender, MouseButtonEventArgs e)
        {
            var messageLabel = sender as Label;
            if (messageLabel == null) return;
            messageLabel.Visibility = Visibility.Hidden;
            messageLabel.Content = string.Empty;
        }

        private void ShowDistinguishedNamesCheckbox_OnChecked(object sender, RoutedEventArgs e)
        {
            foreach (var column in DataGrid.Columns)
            {
                if (column.Header.ToString().Contains("Distinguished Name"))
                {
                    column.Visibility = Visibility.Visible;
                }
            }
        }

        private void ShowDistinguishedNamesCheckbox_OnUnchecked(object sender, RoutedEventArgs e)
        {
            foreach (var column in DataGrid.Columns)
            {
                if (column.Header.ToString().Contains("Distinguished Name"))
                {
                    column.Visibility = Visibility.Hidden;
                }
            }
        }
    }
}
