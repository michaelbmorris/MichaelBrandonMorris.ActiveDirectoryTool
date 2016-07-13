using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Extensions.PrimitiveExtensions;

namespace ActiveDirectoryTool
{
    public partial class View
    {
        private void DataGrid_AutoGeneratingColumn(
            object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            e.Column.Header = e.Column.Header.ToString().SpaceCamelCase();
        }

        private void Label_MouseDoubleClick(
            object sender, MouseButtonEventArgs e)
        {
            var messageLabel = sender as Label;
            if (messageLabel == null) return;
            messageLabel.Visibility = Visibility.Hidden;
            messageLabel.Content = string.Empty;
        }
    }
}