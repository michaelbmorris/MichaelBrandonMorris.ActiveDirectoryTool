using System;
using System.Data;
using System.Windows;

namespace ActiveDirectoryToolWpf
{
    /// <summary>
    /// Interaction logic for ActiveDirectoryToolView.xaml
    /// </summary>
    public partial class ActiveDirectoryToolView : IActiveDirectoryToolView
    {
        public ActiveDirectoryToolView()
        {
            ViewModel = new ActiveDirectoryToolViewModel(this);
            DataContext = new ActiveDirectoryScopeFetcher().Scope;
            InitializeComponent();
            DataGrid.EnableColumnVirtualization = true;
            DataGrid.EnableRowVirtualization = true;
        }

        public event Action GetUsersClicked;
        public event Action GetUsersGroupsClicked;
        public event Action GetDirectReportsClicked;
        public event Action GetComputersClicked;
        public event Action GetGroupsClicked;

        public ActiveDirectoryScope Scope =>
                            TreeView.SelectedItem as ActiveDirectoryScope;

        public ActiveDirectoryToolViewModel ViewModel { get; }

        public void SetDataGridData(DataView dataView)
        {
            DataGrid.ItemsSource = dataView;
        }

        public void ShowMessage(string message)
        {
            MessageBox.Show(message);
        }

        public void ToggleEnabled()
        {
            IsEnabled = !IsEnabled;
        }

        private void GetUsersButton_Click(object sender, RoutedEventArgs e)
        {
            if(Scope != null)
                GetUsersClicked?.Invoke();
        }

        private void GetUsersGroupsButton_Click(object sender, RoutedEventArgs e)
        {
            if(Scope != null)
                GetUsersGroupsClicked?.Invoke();
        }
    }
}