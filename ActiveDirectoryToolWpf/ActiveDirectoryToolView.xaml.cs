using System;
using System.Data;
using System.Windows;

namespace ActiveDirectoryToolWpf
{
    public enum QueryType
    {
        Users,
        Groups,
        UserGroups,
        DirectReports,
        Computers
    }

    public partial class ActiveDirectoryToolView : IActiveDirectoryToolView
    {
        private QueryType _lastQueryType;

        public ActiveDirectoryToolView()
        {
            ViewModel = new ActiveDirectoryToolViewModel(this);
            DataContext = new ActiveDirectoryScopeFetcher().Scope;
            InitializeComponent();
            ProgressBar.Visibility = Visibility.Hidden;
            DataGrid.EnableColumnVirtualization = true;
            DataGrid.EnableRowVirtualization = true;
        }

        public ActiveDirectoryToolViewModel ViewModel { get; }
        public event Action GetComputersClicked;

        public event Action GetDirectReportsClicked;

        public event Action GetGroupsClicked;

        public event Action GetUsersClicked;

        public event Action GetUsersGroupsClicked;

        public ActiveDirectoryScope Scope =>
            TreeView.SelectedItem as ActiveDirectoryScope;

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

        public void ToggleProgressBarVisibility()
        {
            ProgressBar.Visibility =
                ProgressBar.Visibility == Visibility.Visible
                    ? Visibility.Hidden
                    : Visibility.Visible;
        }

        private void GetDirectReports_Click(object sender, RoutedEventArgs e)
        {
            if (Scope != null)
                GetDirectReportsClicked?.Invoke();
            _lastQueryType = QueryType.DirectReports;
        }

        private void GetGroupsButton_Click(object sender, RoutedEventArgs e)
        {
            if (Scope != null)
                GetGroupsClicked?.Invoke();
            _lastQueryType = QueryType.Groups;
        }

        private void GetUsersButton_Click(object sender, RoutedEventArgs e)
        {
            if (Scope != null)
                GetUsersClicked?.Invoke();
            _lastQueryType = QueryType.Users;
        }

        private void GetUsersGroupsButton_Click(object sender,
            RoutedEventArgs e)
        {
            if (Scope != null)
                GetUsersGroupsClicked?.Invoke();
            _lastQueryType = QueryType.UserGroups;
        }

        private void WriteToFile__Click(object sender, RoutedEventArgs e)
        {
            if (DataGrid.Items.Count <= 0) return;
            var fileWriter = new DataFileWriter
            {
                Data = DataGrid,
                Scope = Scope.Context,
                QueryType = _lastQueryType
            };
            fileWriter.WriteToCsv();
        }
    }
}