using System;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using PrimitiveExtensions;

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
        private const string UserDistinguishedName = "UserDistinguishedName";
        private const string GroupDistinguishedName = "GroupDistinguishedName";

        private const string DirectReportDistinguishedName =
            "DirectReportDistinguishedName";

        private const string GetUserGroupsMenuItemHeader =
            "User - Get Groups";

        private const string GetUserDirectReportsMenuItemHeader =
            "User - Get Direct Reports";

        private const string GetDirectReportDirectReportsMenuItemHeader =
            "Direct Report - Get Direct Reports";

        private const string GetDirectReportsUserGroupsMenuItemHeader =
            "Direct Report - Get Groups";

        private const string GetGroupComputersMenuItemHeader =
            "Group - Get Computers";

        private const string GetGroupUsersMenuItemHeader =
            "Group - Get Users";

        private const string NoOrganizationalUnitSelectedErrorMessage =
            "Please select an OU.";

        private const string WroteDataMessage = "Wrote data to ";

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
        public string SelectedItemDistinguishedName { get; set; }

        public ActiveDirectoryScope Scope =>
            TreeView.SelectedItem as ActiveDirectoryScope;

        public event Action GetComputersClicked;
        public event Action GetDirectReportsClicked;
        public event Action GetGroupsClicked;
        public event Action GetUserGroupsClicked;
        public event Action GetUsersClicked;
        public event Action GetUsersGroupsClicked;
        public event Action GetGroupUsersClicked;
        public event Action GetUserDirectReportsClicked;
        public event Action GetGroupComputersClicked;

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

        public void GenerateContextMenu()
        {
            var contextMenu = new ContextMenu();
            if (_lastQueryType == QueryType.Users ||
                _lastQueryType == QueryType.DirectReports ||
                _lastQueryType == QueryType.UserGroups)
            {
                var getUserGroupsMenuItem = new MenuItem
                {
                    Header = GetUserGroupsMenuItemHeader
                };

                getUserGroupsMenuItem.Click += GetUserGroupsMenuItem_Click;
                contextMenu.Items.Add(getUserGroupsMenuItem);
                var getUserDirectReportsMenuItem = new MenuItem
                {
                    Header = GetUserDirectReportsMenuItemHeader
                };

                getUserDirectReportsMenuItem.Click +=
                    GetUserDirectReportsMenuItem_Click;
                contextMenu.Items.Add(getUserDirectReportsMenuItem);
            }

            if (_lastQueryType == QueryType.Groups ||
                _lastQueryType == QueryType.UserGroups)
            {
                var getGroupUsersMenuItem = new MenuItem
                {
                    Header = GetGroupUsersMenuItemHeader
                };

                getGroupUsersMenuItem.Click += GetGroupUsersMenuItem_Click;
                contextMenu.Items.Add(getGroupUsersMenuItem);
                var getGroupComputersMenuItem = new MenuItem
                {
                    Header = GetGroupComputersMenuItemHeader
                };

                getGroupComputersMenuItem.Click +=
                    GetGroupComputersMenuItem_Click;
                contextMenu.Items.Add(getGroupComputersMenuItem);
            }

            if (_lastQueryType == QueryType.DirectReports)
            {
                var getDirectReportsUserGroupsMenuItem = new MenuItem
                {
                    Header = GetDirectReportsUserGroupsMenuItemHeader
                };

                getDirectReportsUserGroupsMenuItem.Click +=
                    GetDirectReportUserGroupsMenuItem_Click;
                contextMenu.Items.Add(getDirectReportsUserGroupsMenuItem);
                var getDirectReportDirectReportsMenuItem = new MenuItem
                {
                    Header = GetDirectReportDirectReportsMenuItemHeader
                };

                getDirectReportDirectReportsMenuItem.Click +=
                    GetDirectReportDirectReportsMenuItem_Click;
                contextMenu.Items.Add(getDirectReportDirectReportsMenuItem);
            }

            DataGrid.ContextMenu = contextMenu;
        }

        private void DataGrid_AutoGeneratingColumn(
            object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            e.Column.Header = e.Column.Header.ToString().SpaceCamelCase();
        }

        private void GetComputers_Click(object sender, RoutedEventArgs e)
        {
            _lastQueryType = QueryType.Computers;
            if (Scope != null)
                GetComputersClicked?.Invoke();
            else
                ShowMessage(NoOrganizationalUnitSelectedErrorMessage);
        }

        private void GetDirectReports_Click(object sender, RoutedEventArgs e)
        {
            _lastQueryType = QueryType.DirectReports;
            if (Scope != null)
                GetDirectReportsClicked?.Invoke();
            else
                ShowMessage(NoOrganizationalUnitSelectedErrorMessage);
        }

        private void GetGroupsButton_Click(object sender, RoutedEventArgs e)
        {
            _lastQueryType = QueryType.Groups;
            if (Scope != null)
                GetGroupsClicked?.Invoke();
            else
                ShowMessage(NoOrganizationalUnitSelectedErrorMessage);
        }

        private void GetUsersButton_Click(object sender, RoutedEventArgs e)
        {
            _lastQueryType = QueryType.Users;
            if (Scope != null)
                GetUsersClicked?.Invoke();
            else
                ShowMessage(NoOrganizationalUnitSelectedErrorMessage);
        }

        private void GetUsersGroupsButton_Click(object sender,
            RoutedEventArgs e)
        {
            _lastQueryType = QueryType.UserGroups;
            if (Scope != null)
                GetUsersGroupsClicked?.Invoke();
            else
                ShowMessage(NoOrganizationalUnitSelectedErrorMessage);
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
            ShowMessage(WroteDataMessage + fileWriter.WriteToCsv());
        }

        private void GetUserGroupsMenuItem_Click(object sender, EventArgs e)
        {
            _lastQueryType = QueryType.UserGroups;
            var row = (DataRowView) DataGrid.SelectedItem;
            SelectedItemDistinguishedName =
                row[UserDistinguishedName].ToString();
            GetUserGroupsClicked?.Invoke();
        }

        private void GetGroupUsersMenuItem_Click(object sender, EventArgs e)
        {
            _lastQueryType = QueryType.Users;
            var row = (DataRowView) DataGrid.SelectedItem;
            SelectedItemDistinguishedName =
                row[GroupDistinguishedName].ToString();
            GetGroupUsersClicked?.Invoke();
        }

        private void GetUserDirectReportsMenuItem_Click(
            object sender, EventArgs e)
        {
            _lastQueryType = QueryType.DirectReports;
            var row = (DataRowView) DataGrid.SelectedItem;
            SelectedItemDistinguishedName =
                row[UserDistinguishedName].ToString();
            GetUserDirectReportsClicked?.Invoke();
        }

        private void GetDirectReportDirectReportsMenuItem_Click(
            object sender, EventArgs e)
        {
            _lastQueryType = QueryType.DirectReports;
            var row = (DataRowView) DataGrid.SelectedItem;
            SelectedItemDistinguishedName =
                row[DirectReportDistinguishedName].ToString();
            GetUserDirectReportsClicked?.Invoke();
        }

        private void GetGroupComputersMenuItem_Click(
            object sender, EventArgs e)
        {
            _lastQueryType = QueryType.Computers;
            var row = (DataRowView) DataGrid.SelectedItem;
            SelectedItemDistinguishedName =
                row[GroupDistinguishedName].ToString();
            GetGroupComputersClicked?.Invoke();
        }

        private void GetDirectReportUserGroupsMenuItem_Click(
            object sender, EventArgs e)
        {
            _lastQueryType = QueryType.UserGroups;
            var row = (DataRowView) DataGrid.SelectedItem;
            SelectedItemDistinguishedName =
                row[DirectReportDistinguishedName].ToString();
            GetUserGroupsClicked?.Invoke();
        }
    }
}