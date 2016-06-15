using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ActiveDirectoryToolWpf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ActiveDirectorySearcher searcher;
        private ActiveDirectoryScope selectedScope;
        private DataPreparer dataPreparer;

        public MainWindow()
        {
            DataContext = new ActiveDirectoryScopeFetcher().Scope;
            InitializeComponent();
            selectedScope = null;
            DataGrid.EnableColumnVirtualization = true;
            DataGrid.EnableRowVirtualization = true;
        }

        private void GetUsersButton_Click(object sender, RoutedEventArgs e)
        {
            selectedScope = TreeView.SelectedItem as ActiveDirectoryScope;
            Debug.WriteLine(selectedScope);
            if (selectedScope != null)
            {
                searcher = new ActiveDirectorySearcher
                {
                    Scope = selectedScope
                };
                dataPreparer = new DataPreparer
                {
                    Principals = searcher.GetUsers()
                };
                DataGrid.ItemsSource = dataPreparer.GetResults().ToDataTable().AsDataView();
            }
            else
            {
                MessageBox.Show("Please select a scope.");
            }
        }
    }
}
