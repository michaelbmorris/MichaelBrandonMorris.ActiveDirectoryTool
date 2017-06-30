using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MichaelBrandonMorris.Extensions.PrimitiveExtensions;

namespace MichaelBrandonMorris.ActiveDirectoryTool
{
    /// <summary>
    ///     Class MainWindow.
    /// </summary>
    /// <seealso cref="Window" />
    /// <seealso cref="System.Windows.Markup.IComponentConnector" />
    /// TODO Edit XML Comment Template for MainWindow
    public partial class MainWindow
    {
        /// <summary>
        ///     Handles the AutoGeneratingColumn event of the DataGrid
        ///     control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">
        ///     The
        ///     <see cref="DataGridAutoGeneratingColumnEventArgs" />
        ///     instance containing the event data.
        /// </param>
        /// TODO Edit XML Comment Template for DataGrid_AutoGeneratingColumn
        private void DataGrid_AutoGeneratingColumn(
            object sender,
            DataGridAutoGeneratingColumnEventArgs e)
        {
            e.Column.Header = e.Column.Header.ToString().SpaceCamelCase();

            if (ShowDistinguishedNamesCheckbox.IsChecked != null
                && ShowDistinguishedNamesCheckbox.IsChecked.Value)
            {
                return;
            }

            if (e.Column.Header.ToString().Contains("Distinguished Name"))
            {
                e.Column.Visibility = Visibility.Hidden;
            }
        }

        /// <summary>
        ///     Handles the MouseDoubleClick event of the Label
        ///     control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">
        ///     The <see cref="MouseButtonEventArgs" />
        ///     instance containing the event data.
        /// </param>
        /// TODO Edit XML Comment Template for Label_MouseDoubleClick
        private void Label_MouseDoubleClick(
            object sender,
            MouseButtonEventArgs e)
        {
            var messageLabel = sender as Label;

            if (messageLabel == null)
            {
                return;
            }

            messageLabel.Visibility = Visibility.Hidden;
            messageLabel.Content = string.Empty;
        }

        /// <summary>
        ///     Handles the OnChecked event of the
        ///     ShowDistinguishedNamesCheckbox control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">
        ///     The <see cref="RoutedEventArgs" /> instance
        ///     containing the event data.
        /// </param>
        /// TODO Edit XML Comment Template for ShowDistinguishedNamesCheckbox_OnChecked
        private void ShowDistinguishedNamesCheckbox_OnChecked(
            object sender,
            RoutedEventArgs e)
        {
            foreach (var column in DataGrid.Columns)
            {
                if (column.Header.ToString().Contains("Distinguished Name"))
                {
                    column.Visibility = Visibility.Visible;
                }
            }
        }

        /// <summary>
        ///     Handles the OnUnchecked event of the
        ///     ShowDistinguishedNamesCheckbox control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">
        ///     The <see cref="RoutedEventArgs" /> instance
        ///     containing the event data.
        /// </param>
        /// TODO Edit XML Comment Template for ShowDistinguishedNamesCheckbox_OnUnchecked
        private void ShowDistinguishedNamesCheckbox_OnUnchecked(
            object sender,
            RoutedEventArgs e)
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