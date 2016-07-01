using System;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using PrimitiveExtensions;

namespace ActiveDirectoryToolWpf
{
    public enum QueryType
    {
        Any,
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

        public QueryType QueryType { get; set; }

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

        public string SearchText { get; set; }

        public event Action GetComputersButtonClicked;
        public event Action GetDirectReportsButtonClicked;
        public event Action GetGroupsButtonClicked;
        public event Action GetUserGroupsMenuItemClicked;
        public event Action GetUsersButtonClicked;
        public event Action GetUsersGroupsButtonClicked;
        public event Action GetGroupUsersMenuItemClicked;
        public event Action GetUserDirectReportsMenuItemClicked;
        public event Action GetGroupComputersMenuItemClicked;
        public event Action SearchButtonClicked;

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
            if (QueryType == QueryType.Users ||
                QueryType == QueryType.DirectReports ||
                QueryType == QueryType.UserGroups)
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

            if (QueryType == QueryType.Groups ||
                QueryType == QueryType.UserGroups)
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

            if (QueryType == QueryType.DirectReports)
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

        private void GetComputersButton_OnClick(object sender, RoutedEventArgs e)
        {
            QueryType = QueryType.Computers;
            if (Scope != null)
                GetComputersButtonClicked?.Invoke();
            else
                ShowMessage(NoOrganizationalUnitSelectedErrorMessage);
        }

        private void GetDirectReportsButton_OnClick(object sender, RoutedEventArgs e)
        {
            QueryType = QueryType.DirectReports;
            if (Scope != null)
                GetDirectReportsButtonClicked?.Invoke();
            else
                ShowMessage(NoOrganizationalUnitSelectedErrorMessage);
        }

        private void GetGroupsButton_OnClick(object sender, RoutedEventArgs e)
        {
            QueryType = QueryType.Groups;
            if (Scope != null)
                GetGroupsButtonClicked?.Invoke();
            else
                ShowMessage(NoOrganizationalUnitSelectedErrorMessage);
        }

        private void GetUsersButton_OnClick(object sender, RoutedEventArgs e)
        {
            QueryType = QueryType.Users;
            if (Scope != null)
                GetUsersButtonClicked?.Invoke();
            else
                ShowMessage(NoOrganizationalUnitSelectedErrorMessage);
        }

        private void GetUsersGroupsButton_OnClick(object sender,
            RoutedEventArgs e)
        {
            QueryType = QueryType.UserGroups;
            if (Scope != null)
                GetUsersGroupsButtonClicked?.Invoke();
            else
                ShowMessage(NoOrganizationalUnitSelectedErrorMessage);
        }

        private void WriteToFileButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (DataGrid.Items.Count <= 0) return;
            var fileWriter = new DataFileWriter
            {
                Data = DataGrid,
                Scope = Scope.Context,
                QueryType = QueryType
            };
            ShowMessage(WroteDataMessage + fileWriter.WriteToCsv());
        }

        private void GetUserGroupsMenuItem_Click(object sender, EventArgs e)
        {
            QueryType = QueryType.UserGroups;
            var row = (DataRowView) DataGrid.SelectedItem;
            SelectedItemDistinguishedName =
                row[UserDistinguishedName].ToString();
            GetUserGroupsMenuItemClicked?.Invoke();
        }

        private void GetGroupUsersMenuItem_Click(object sender, EventArgs e)
        {
            QueryType = QueryType.Users;
            var row = (DataRowView) DataGrid.SelectedItem;
            SelectedItemDistinguishedName =
                row[GroupDistinguishedName].ToString();
            GetGroupUsersMenuItemClicked?.Invoke();
        }

        private void GetUserDirectReportsMenuItem_Click(
            object sender, EventArgs e)
        {
            QueryType = QueryType.DirectReports;
            var row = (DataRowView) DataGrid.SelectedItem;
            SelectedItemDistinguishedName =
                row[UserDistinguishedName].ToString();
            GetUserDirectReportsMenuItemClicked?.Invoke();
        }

        private void GetDirectReportDirectReportsMenuItem_Click(
            object sender, EventArgs e)
        {
            QueryType = QueryType.DirectReports;
            var row = (DataRowView) DataGrid.SelectedItem;
            SelectedItemDistinguishedName =
                row[DirectReportDistinguishedName].ToString();
            GetUserDirectReportsMenuItemClicked?.Invoke();
        }

        private void GetGroupComputersMenuItem_Click(
            object sender, EventArgs e)
        {
            QueryType = QueryType.Computers;
            var row = (DataRowView) DataGrid.SelectedItem;
            SelectedItemDistinguishedName =
                row[GroupDistinguishedName].ToString();
            GetGroupComputersMenuItemClicked?.Invoke();
        }

        private void GetDirectReportUserGroupsMenuItem_Click(
            object sender, EventArgs e)
        {
            QueryType = QueryType.UserGroups;
            var row = (DataRowView) DataGrid.SelectedItem;
            SelectedItemDistinguishedName =
                row[DirectReportDistinguishedName].ToString();
            GetUserGroupsMenuItemClicked?.Invoke();
        }

        /*private void SearchButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (GroupRadioButton.IsChecked != null &&
                GroupRadioButton.IsChecked.Value)
            {
                QueryType = QueryType.Groups;
            }
            else if (UserRadioButton.IsChecked != null &&
                UserRadioButton.IsChecked.Value)
            {
                QueryType = QueryType.Users;
            }
            else if (ComputerRadioButton.IsChecked != null &&
                ComputerRadioButton.IsChecked.Value)
            {
                QueryType = QueryType.Computers;
            }
            else
            {
                QueryType = QueryType.Any;
            }
            if(Scope == null)
                ShowMessage(NoOrganizationalUnitSelectedErrorMessage);
            else if(SearchText.IsNullOrWhiteSpace())
                ShowMessage("Please enter a search query.");
            else
                SearchButtonClicked?.Invoke();

        }*/

        /*private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            SearchText = SearchTextBox.Text;
        }*/
    }
}