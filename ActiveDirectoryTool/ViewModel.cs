using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Deployment.Application;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Extensions.CollectionExtensions;
using Extensions.PrimitiveExtensions;
using GalaSoft.MvvmLight.CommandWpf;
using static ActiveDirectoryTool.QueryType;
using static System.Deployment.Application.ApplicationDeployment;

namespace ActiveDirectoryTool
{
    internal class ViewModel : INotifyPropertyChanged
    {
        private const string ComputerDistinguishedName =
            "ComputerDistinguishedName";

        private const string ContainerGroupDistinguishedName =
            "ContainerGroupDistinguishedName";

        private const string ContainerGroupManagedByDistinguishedName =
            "ContainerGroupManagedByDistinguishedName";

        private const string DirectReportDistinguishedName =
            "DirectReportDistinguishedName";

        private const string GroupDistinguishedName = "GroupDistinguishedName";

        private const string GroupManagedByDistinguishedName =
            "GroupManagedByDistinguishedName";

        private const string HelpFile = "ActiveDirectoryToolHelp.chm";

        private const string UserDistinguishedName = "UserDistinguishedName";

        private const string UserManagerDistinguishedName =
            "UserManagerDistinguishedName";

        private readonly MenuItem _computerGetGroupsMenuItem;
        private readonly MenuItem _computerGetSummaryMenuItem;
        private readonly MenuItem _containerGroupGetComputersMenuItem;

        private readonly MenuItem
            _containerGroupGetManagedByDirectReportsMenuItem;

        private readonly MenuItem _containerGroupGetManagedByGroupsMenuItem;
        private readonly MenuItem _containerGroupGetManagedBySummaryMenuItem;
        private readonly MenuItem _containerGroupGetSummaryMenuItem;
        private readonly MenuItem _containerGroupGetUsersDirectReportsMenuItem;
        private readonly MenuItem _containerGroupGetUsersGroupsMenuItem;
        private readonly MenuItem _containerGroupGetUsersMenuItem;
        private readonly MenuItem _directReportGetDirectReportsMenuItem;
        private readonly MenuItem _directReportGetGroupsMenuItem;
        private readonly MenuItem _directReportGetSummaryMenuItem;
        private readonly MenuItem _groupGetComputersMenuItem;
        private readonly MenuItem _groupGetManagedByDirectReportsMenuItem;
        private readonly MenuItem _groupGetManagedByGroupsMenuItem;
        private readonly MenuItem _groupGetManagedBySummaryMenuItem;
        private readonly MenuItem _groupGetSummaryMenuItem;
        private readonly MenuItem _groupGetUsersDirectReportsMenuItem;
        private readonly MenuItem _groupGetUsersGroupsMenuItem;
        private readonly MenuItem _groupGetUsersMenuItem;
        private readonly MenuItem _managerGetDirectReportsMenuItem;
        private readonly MenuItem _managerGetGroupsMenuItem;
        private readonly MenuItem _managerGetSummaryMenuItem;
        private readonly MenuItem _userGetDirectReportsMenuItem;
        private readonly MenuItem _userGetGroupsMenuItem;
        private readonly MenuItem _userGetSummaryMenuItem;

        private AboutWindow _aboutWindow;
        private Visibility _cancelButtonVisibility;
        private bool _computerSearchIsChecked;
        private List<MenuItem> _contextMenuItems;
        private Visibility _contextMenuVisibility;
        private DataView _data;
        private bool _groupSearchIsChecked;
        private string _messageContent;
        private Visibility _messageVisibility;
        private Visibility _progressBarVisibility;
        private ObservableCollection<Query> _queries;
        private string _searchText;
        private bool _showDistinguishedNames;
        private bool _userSearchIsChecked;
        private bool _viewIsEnabled;

        public ViewModel()
        {
            SetViewVariables();
            try
            {
                RootScope = new ScopeFetcher().Scope;
            }
            catch (PrincipalServerDownException)
            {
                ShowMessage(
                    "The Active Directory server could not be contacted.");
            }
            Queries = new ObservableCollection<Query>();
            _computerGetGroupsMenuItem = new MenuItem
            {
                Header = "Computer - Get Groups",
                Command = GetComputerGroupsCommand
            };
            _computerGetSummaryMenuItem = new MenuItem
            {
                Header = "Computer - Get Summary",
                Command = GetComputerSummaryCommand
            };
            _containerGroupGetComputersMenuItem = new MenuItem
            {
                Header = "Container Group - Get Computers",
                Command = GetContainerGroupComputersCommand
            };
            _containerGroupGetManagedByDirectReportsMenuItem = new MenuItem
            {
                Header = "Container Group - Get Managed By's Direct Reports",
                Command = GetContainerGroupManagedByDirectReportsCommand
            };
            _containerGroupGetManagedByGroupsMenuItem = new MenuItem
            {
                Header = "Container Group - Get Managed By's Groups",
                Command = GetContainerGroupManagedByGroupsCommand
            };
            _containerGroupGetManagedBySummaryMenuItem = new MenuItem
            {
                Header = "Container Group - Get Managed By Summary",
                Command = GetContainerGroupManagedBySummaryCommand
            };
            _containerGroupGetSummaryMenuItem = new MenuItem
            {
                Header = "Container Group - Get Summary",
                Command = GetContainerGroupSummaryCommand
            };
            _containerGroupGetUsersMenuItem = new MenuItem
            {
                Header = "Container Group - Get Users",
                Command = GetContainerGroupUsersCommand
            };
            _containerGroupGetUsersDirectReportsMenuItem = new MenuItem
            {
                Header = "Container Group - Get Users' Direct Reports",
                Command = GetContainerGroupUsersDirectReportsCommand
            };
            _containerGroupGetUsersGroupsMenuItem = new MenuItem
            {
                Header = "Container Group - Get Users' Groups",
                Command = GetContainerGroupUsersGroupsCommand
            };
            _directReportGetDirectReportsMenuItem = new MenuItem
            {
                Header = "Direct Report - Get Direct Reports",
                Command = GetDirectReportDirectReportsCommand
            };
            _directReportGetGroupsMenuItem = new MenuItem
            {
                Header = "Direct Report - Get Groups",
                Command = GetDirectReportGroupsCommand
            };
            _directReportGetSummaryMenuItem = new MenuItem
            {
                Header = "Direct Report - Get Summary",
                Command = GetDirectReportSummaryCommand
            };
            _groupGetComputersMenuItem = new MenuItem
            {
                Header = "Group - Get Computers",
                Command = GetGroupComputersCommand
            };
            _groupGetUsersMenuItem = new MenuItem
            {
                Header = "Group - Get Users",
                Command = GetGroupUsersCommand
            };
            _groupGetUsersDirectReportsMenuItem = new MenuItem
            {
                Header = "Group - Get Users' Direct Reports",
                Command = GetGroupUsersDirectReportsCommand
            };
            _groupGetUsersGroupsMenuItem = new MenuItem
            {
                Header = "Group - Get Users' Groups",
                Command = GetGroupUsersGroupsCommand
            };
            _groupGetSummaryMenuItem = new MenuItem
            {
                Header = "Group - Get Summary",
                Command = GetGroupSummaryCommand
            };
            _managerGetDirectReportsMenuItem = new MenuItem
            {
                Header = "Manager - Get Direct Reports",
                Command = GetManagerDirectReportsCommand
            };
            _managerGetGroupsMenuItem = new MenuItem
            {
                Header = "Manager - Get Groups",
                Command = GetManagerGroupsCommand
            };
            _managerGetSummaryMenuItem = new MenuItem
            {
                Header = "Manager - Get Summary",
                Command = GetManagerSummaryCommand
            };
            _userGetDirectReportsMenuItem = new MenuItem
            {
                Header = "User - Get Direct Reports",
                Command = GetUserDirectReportsCommand
            };
            _userGetGroupsMenuItem = new MenuItem
            {
                Header = "User - Get Groups",
                Command = GetUserGroupsCommand
            };
            _userGetSummaryMenuItem = new MenuItem
            {
                Header = "User - Get Summary",
                Command = GetUserSummaryCommand
            };
        }

        public Visibility CancelButtonVisibility
        {
            get
            {
                return _cancelButtonVisibility;
            }
            set
            {
                _cancelButtonVisibility = value;
                NotifyPropertyChanged();
            }
        }

        public ICommand CancelCommand { get; private set; }

        public bool ComputerSearchIsChecked
        {
            get
            {
                return _computerSearchIsChecked;
            }
            set
            {
                _computerSearchIsChecked = value;
                NotifyPropertyChanged();
            }
        }

        public List<MenuItem> ContextMenuItems
        {
            get
            {
                return _contextMenuItems;
            }
            private set
            {
                _contextMenuItems = value;
                ContextMenuVisibility = _contextMenuItems.IsNullOrEmpty()
                    ? Visibility.Hidden
                    : Visibility.Visible;
                NotifyPropertyChanged();
            }
        }

        public Visibility ContextMenuVisibility
        {
            get
            {
                return _contextMenuVisibility;
            }
            set
            {
                _contextMenuVisibility = value;
                NotifyPropertyChanged();
            }
        }

        public Scope CurrentScope { get; set; }

        public DataView Data
        {
            get
            {
                return _data;
            }
            private set
            {
                _data = value;
                NotifyPropertyChanged();
            }
        }

        public ICommand GetOuComputersCommand { get; private set; }
        public ICommand GetOuGroupsCommand { get; private set; }
        public ICommand GetOuGroupsUsersCommand { get; private set; }
        public ICommand GetOuUsersCommand { get; private set; }
        public ICommand GetOuUsersDirectReportsCommand { get; private set; }
        public ICommand GetOuUsersGroupsCommand { get; private set; }

        public bool GroupSearchIsChecked
        {
            get
            {
                return _groupSearchIsChecked;
            }
            set
            {
                _groupSearchIsChecked = value;
                NotifyPropertyChanged();
            }
        }

        public string MessageContent
        {
            get
            {
                return _messageContent;
            }
            set
            {
                _messageContent = value;
                NotifyPropertyChanged();
            }
        }

        public Visibility MessageVisibility
        {
            get
            {
                return _messageVisibility;
            }
            set
            {
                _messageVisibility = value;
                NotifyPropertyChanged();
            }
        }

        public ICommand OpenAboutWindowCommand { get; private set; }
        public ICommand OpenHelpWindowCommand { get; private set; }
        public ICommand PreviousQueryCommand { get; private set; }

        public Visibility ProgressBarVisibility
        {
            get
            {
                return _progressBarVisibility;
            }
            set
            {
                _progressBarVisibility = value;
                NotifyPropertyChanged();
            }
        }

        public ObservableCollection<Query> Queries
        {
            get
            {
                return _queries;
            }
            set
            {
                _queries = value;
                NotifyPropertyChanged();
            }
        }

        public Scope RootScope { get; }
        public ICommand SearchCommand { get; private set; }
        public ICommand SearchOuCommand { get; private set; }

        public string SearchText
        {
            get
            {
                return _searchText;
            }
            set
            {
                _searchText = value;
                NotifyPropertyChanged();
            }
        }

        public ICommand SelectionChangedCommand { get; private set; }

        public bool ShowDistinguishedNames
        {
            get
            {
                return _showDistinguishedNames;
            }
            set
            {
                _showDistinguishedNames = value;
                NotifyPropertyChanged();
            }
        }

        public bool UserSearchIsChecked
        {
            get
            {
                return _userSearchIsChecked;
            }
            set
            {
                _userSearchIsChecked = value;
                NotifyPropertyChanged();
            }
        }

        public string Version { get; private set; }

        public bool ViewIsEnabled
        {
            get
            {
                return _viewIsEnabled;
            }
            private set
            {
                _viewIsEnabled = value;
                NotifyPropertyChanged();
            }
        }

        public ICommand WriteToFileCommand { get; private set; }
        private ICommand GetComputerGroupsCommand { get; set; }
        private ICommand GetComputerSummaryCommand { get; set; }
        private ICommand GetContainerGroupComputersCommand { get; set; }

        private ICommand GetContainerGroupManagedByDirectReportsCommand { get;
            set; }

        private ICommand GetContainerGroupManagedByGroupsCommand { get; set; }

        private ICommand GetContainerGroupManagedBySummaryCommand { get; set; }

        private ICommand GetContainerGroupSummaryCommand { get; set; }
        private ICommand GetContainerGroupUsersCommand { get; set; }

        private ICommand GetContainerGroupUsersDirectReportsCommand { get; set;
        }

        private ICommand GetContainerGroupUsersGroupsCommand { get; set; }
        private ICommand GetDirectReportDirectReportsCommand { get; set; }
        private ICommand GetDirectReportGroupsCommand { get; set; }
        private ICommand GetDirectReportSummaryCommand { get; set; }
        private ICommand GetGroupComputersCommand { get; set; }
        private ICommand GetGroupSummaryCommand { get; set; }
        private ICommand GetGroupUsersCommand { get; set; }
        private ICommand GetGroupUsersDirectReportsCommand { get; set; }
        private ICommand GetGroupUsersGroupsCommand { get; set; }
        private ICommand GetManagerDirectReportsCommand { get; set; }
        private ICommand GetManagerGroupsCommand { get; set; }
        private ICommand GetManagerSummaryCommand { get; set; }
        private ICommand GetUserDirectReportsCommand { get; set; }
        private ICommand GetUserGroupsCommand { get; set; }
        private ICommand GetUserSummaryCommand { get; set; }
        private IList SelectedDataRowViews { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        private static string GetComputerDistinguishedName(
            DataRowView dataRowView)
        {
            return dataRowView[ComputerDistinguishedName].ToString();
        }

        private static string GetContainerGroupDistinguishedName(
            DataRowView dataRowView)
        {
            return dataRowView[ContainerGroupDistinguishedName].ToString();
        }

        private static string GetContainerGroupManagedByDistinguishedName(
            DataRowView dataRowView)
        {
            return dataRowView[ContainerGroupManagedByDistinguishedName]
                .ToString();
        }

        private static string GetDirectReportDistinguishedName(
            DataRowView dataRowView)
        {
            return dataRowView[DirectReportDistinguishedName].ToString();
        }

        private static string GetGroupDistinguishedName(
            DataRowView dataRowView)
        {
            return dataRowView[GroupDistinguishedName].ToString();
        }

        private static string GetManagerDistinguishedName(
            DataRowView dataRowView)
        {
            return dataRowView[UserManagerDistinguishedName].ToString();
        }

        private static string GetUserDistinguishedName(DataRowView dataRowView)
        {
            return dataRowView[UserDistinguishedName].ToString();
        }

        private static void OpenHelpWindowCommandExecute() => Process.Start(
            HelpFile);

        private void CancelCommandExecute()
        {
            Queries.Peek()?.Cancel();
            ResetQuery();
        }

        private void FinishTask()
        {
            ProgressBarVisibility = Visibility.Hidden;
            CancelButtonVisibility = Visibility.Hidden;
            ViewIsEnabled = true;
        }

        private IEnumerable<MenuItem> GenerateComputerContextMenuItems()
        {
            var computerContextMenuItems = new List<MenuItem>();
            var queryType = Queries.Peek().QueryType;
            if (queryType != ComputersGroups)
                computerContextMenuItems.Add(_computerGetGroupsMenuItem);
            if (queryType != ComputersSummaries)
                computerContextMenuItems.Add(_computerGetSummaryMenuItem);
            return computerContextMenuItems;
        }

        private List<MenuItem> GenerateContextMenuItems()
        {
            var contextMenuItems = new List<MenuItem>();
            if (Queries.Peek() == null) return null;
            if (Data.Table.Columns.Contains(ComputerDistinguishedName))
            {
                contextMenuItems.AddRange(GenerateComputerContextMenuItems());
            }
            if (Data.Table.Columns.Contains(ContainerGroupDistinguishedName))
            {
                contextMenuItems.Add(_containerGroupGetComputersMenuItem);
                contextMenuItems.Add(_containerGroupGetSummaryMenuItem);
                contextMenuItems.Add(_containerGroupGetUsersMenuItem);
                contextMenuItems.Add(_containerGroupGetUsersGroupsMenuItem);
                contextMenuItems.Add(
                    _containerGroupGetUsersDirectReportsMenuItem);
            }
            if (Data.Table.Columns.Contains(
                ContainerGroupManagedByDistinguishedName))
            {
                contextMenuItems.Add(
                    _containerGroupGetManagedByDirectReportsMenuItem);
                contextMenuItems.Add(
                    _containerGroupGetManagedByGroupsMenuItem);
                contextMenuItems.Add(
                    _containerGroupGetManagedBySummaryMenuItem);
            }
            if (Data.Table.Columns.Contains(DirectReportDistinguishedName))
            {
                contextMenuItems.Add(_directReportGetDirectReportsMenuItem);
                contextMenuItems.Add(_directReportGetGroupsMenuItem);
                contextMenuItems.Add(_directReportGetSummaryMenuItem);
            }
            if (Data.Table.Columns.Contains(GroupDistinguishedName))
            {
                contextMenuItems.AddRange(GenerateGroupContextMenuItems());
            }
            if (Data.Table.Columns.Contains(UserManagerDistinguishedName))
            {
                contextMenuItems.Add(_managerGetDirectReportsMenuItem);
                contextMenuItems.Add(_managerGetGroupsMenuItem);
                contextMenuItems.Add(_managerGetSummaryMenuItem);
            }
            if (!Data.Table.Columns.Contains(UserDistinguishedName))
                return contextMenuItems;
            contextMenuItems.AddRange(GenerateUserContextMenuItems());
            return contextMenuItems;
        }

        private IEnumerable<MenuItem> GenerateGroupContextMenuItems()
        {
            var groupContextMenuItems = new List<MenuItem>();
            var queryType = Queries.Peek().QueryType;
            if (queryType != GroupsComputers)
                groupContextMenuItems.Add(_groupGetComputersMenuItem);
            if (queryType != GroupsSummaries)
                groupContextMenuItems.Add(_groupGetSummaryMenuItem);
            if (queryType != GroupsUsers)
                groupContextMenuItems.Add(_groupGetUsersMenuItem);
            if (queryType != GroupsUsersDirectReports)
                groupContextMenuItems.Add(_groupGetUsersDirectReportsMenuItem);
            if (queryType != GroupsUsersGroups)
                groupContextMenuItems.Add(_groupGetUsersGroupsMenuItem);
            return groupContextMenuItems;
        }

        private IEnumerable<MenuItem> GenerateUserContextMenuItems()
        {
            var userContextMenuItems = new List<MenuItem>();
            var queryType = Queries.Peek().QueryType;
            if (queryType != UsersDirectReports)
                userContextMenuItems.Add(_userGetDirectReportsMenuItem);
            if (queryType != UsersGroups)
                userContextMenuItems.Add(_userGetGroupsMenuItem);
            if (queryType != UsersSummaries)
                userContextMenuItems.Add(_userGetSummaryMenuItem);
            return userContextMenuItems;
        }

        private IEnumerable<string> GetComputersDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>().Select(
                GetComputerDistinguishedName).Where(
                    c => !c.IsNullOrWhiteSpace());
        }

        private async void GetContainerGroupComputersCommandExecute()
        {
            await RunQuery(
                GroupsComputers,
                CurrentScope,
                GetContainerGroupsDistinguishedNames());
        }

        private async void
            GetContainerGroupManagedByDirectReportsCommandExecute()
        {
            await RunQuery(
                UsersDirectReports,
                CurrentScope,
                GetContainerGroupsManagedByDistinguishedNames());
        }

        private async void GetContainerGroupManagedByGroupsCommandExecute()
        {
            await RunQuery(
                UsersGroups,
                CurrentScope,
                GetContainerGroupsManagedByDistinguishedNames());
        }

        private async void GetContainerGroupManagedBySummaryCommandExecute()
        {
            await RunQuery(
                UsersSummaries,
                CurrentScope,
                GetContainerGroupsManagedByDistinguishedNames());
        }

        private IEnumerable<string> GetContainerGroupsDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>()
                .Select(GetContainerGroupDistinguishedName)
                .Where(c => !c.IsNullOrWhiteSpace());
        }

        private IEnumerable<string>
            GetContainerGroupsManagedByDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>().Select(
                GetContainerGroupManagedByDistinguishedName).Where(
                    s => !s.IsNullOrWhiteSpace());
        }

        private async void GetContainerGroupSummaryCommandExecute()
        {
            await RunQuery(
                GroupsSummaries,
                CurrentScope,
                GetContainerGroupsDistinguishedNames());
        }

        private async void GetContainerGroupUsersCommandExecute()
        {
            await RunQuery(
                GroupsUsers,
                CurrentScope,
                GetContainerGroupsDistinguishedNames());
        }

        private async void GetContextComputerGroupsCommandExecute()
        {
            await RunQuery(
                ComputersGroups,
                CurrentScope,
                GetComputersDistinguishedNames());
        }

        private async void GetContainerGroupUsersDirectReportsCommandExecute()
        {
            await RunQuery(
                GroupsUsersDirectReports,
                CurrentScope,
                GetContainerGroupsDistinguishedNames());
        }

        private async void GetContextComputerSummaryCommandExecute()
        {
            await RunQuery(
                ComputersSummaries,
                CurrentScope,
                GetComputersDistinguishedNames());
        }

        private async void GetContextDirectReportDirectReportsCommandExecute()
        {
            await RunQuery(
                UsersDirectReports,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        private async void GetContextDirectReportGroupsCommandExecute()
        {
            await RunQuery(
                UsersGroups,
                CurrentScope,
                GetDirectReportsDistinguishedNames());
        }

        private async void GetContextDirectReportSummaryCommandExecute()
        {
            await RunQuery(
                UsersSummaries,
                CurrentScope,
                GetDirectReportsDistinguishedNames());
        }

        private async void GetContextGroupComputersCommandExecute()
        {
            await RunQuery(
                GroupsComputers,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        private async void GetContextGroupSummaryCommandExecute()
        {
            await RunQuery(
                GroupsSummaries,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        private async void GetContextGroupUsersCommandExecute()
        {
            await RunQuery(
                GroupsUsers,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        private async void GetContextGroupUsersDirectReportsCommandExecute()
        {
            await RunQuery(
                GroupsUsersDirectReports,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        private async void GetContextGroupUsersGroupsCommandExecute()
        {
            await RunQuery(
                GroupsUsersGroups,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        private async void GetContextUserDirectReportsCommandExecute()
        {
            await RunQuery(
                UsersDirectReports,
                CurrentScope,
                GetUsersDistinguishedNames());
        }

        private async void GetContextUserGroupsCommandExecute()
        {
            await RunQuery(
                UsersGroups,
                CurrentScope,
                GetUsersDistinguishedNames());
        }

        private async void GetContextUserSummaryCommandExecute()
        {
            await RunQuery(
                UsersSummaries,
                CurrentScope,
                GetUsersDistinguishedNames());
        }

        private IEnumerable<string> GetDirectReportsDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>().Select(
                GetDirectReportDistinguishedName).Where(
                    d => !d.IsNullOrWhiteSpace());
        }

        private IEnumerable<string> GetGroupsDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>().Select(
                GetGroupDistinguishedName).Where(g => !g.IsNullOrWhiteSpace());
        }

        private async void GetManagerDirectReportsCommandExecute()
        {
            await RunQuery(
                UsersDirectReports,
                CurrentScope,
                GetManagersDistinguishedNames());
        }

        private async void GetManagerGroupsCommandExecute()
        {
            await RunQuery(
                UsersGroups,
                CurrentScope,
                GetManagersDistinguishedNames());
        }

        private IEnumerable<string> GetManagersDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>().Select(
                GetManagerDistinguishedName).Where(
                    m => !m.IsNullOrWhiteSpace());
        }

        private async void GetManagerSummaryCommandExecute()
        {
            await RunQuery(
                UsersSummaries,
                CurrentScope,
                GetManagersDistinguishedNames());
        }

        private async void GetOuComputersCommandExecute()
        {
            await RunQuery(OuComputers, CurrentScope);
        }

        private async void GetOuGroupsCommandExecute()
        {
            await RunQuery(OuGroups, CurrentScope);
        }

        private async void GetOuGroupsUsersCommandExecute()
        {
            await RunQuery(OuGroupsUsers, CurrentScope);
        }

        private async void GetOuUsersCommandExecute()
        {
            await RunQuery(OuUsers, CurrentScope);
        }

        private async void GetOuUsersDirectReportsCommandExecute()
        {
            await RunQuery(OuUsersDirectReports, CurrentScope);
        }

        private async void GetOuUsersGroupsCommandExecute()
        {
            await RunQuery(OuUsersGroups, CurrentScope);
        }

        private QueryType GetSearchQueryType()
        {
            if (ComputerSearchIsChecked)
                return SearchComputer;
            if (GroupSearchIsChecked)
                return SearchGroup;
            return UserSearchIsChecked ? SearchUser : None;
        }

        private IEnumerable<string> GetUsersDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>().Select(
                GetUserDistinguishedName).Where(u => !u.IsNullOrWhiteSpace());
        }

        private void HideMessage()
        {
            MessageContent = string.Empty;
            MessageVisibility = Visibility.Hidden;
        }

        private void NotifyPropertyChanged(
            [CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(
                this,
                new PropertyChangedEventArgs(propertyName));
        }

        private void OpenAboutWindowCommandExecute()
        {
            if (_aboutWindow == null || !_aboutWindow.IsVisible)
            {
                _aboutWindow = new AboutWindow();
                _aboutWindow.Show();
            }
            else
                _aboutWindow.Activate();
        }

        private bool OuCommandCanExecute()
        {
            return CurrentScope != null;
        }

        private bool PreviousQueryCommandCanExecute()
        {
            return Queries.Multiple();
        }

        private async void PreviousQueryCommandExecute()
        {
            Queries.Pop();
            await RunQuery(Queries.Pop());
        }

        private void ResetQuery()
        {
            if (Queries.Any())
                Queries.Pop();
        }

        private async Task RunQuery(
            QueryType queryType,
            Scope scope = null,
            IEnumerable<string> selectedItemDistinguishedNames = null,
            string searchText = null)
        {
            await RunQuery(
                new Query(
                    queryType,
                    scope,
                    selectedItemDistinguishedNames,
                    searchText));
        }

        private async Task RunQuery(Query query)
        {
            StartTask();
            try
            {
                Queries.Push(query);
                await Queries.Peek().Execute();
                Data = Queries.Peek().Data.ToDataTable().AsDataView();
                ContextMenuItems = GenerateContextMenuItems();
            }
            catch (NullReferenceException e)
            {
                ShowMessage(e.Message);
                ResetQuery();
            }
            catch (OperationCanceledException)
            {
                ShowMessage("Operation was cancelled.");
                ResetQuery();
            }
            catch (ArgumentNullException)
            {
                ShowMessage(
                    "No results of desired type found in selected context.");
                ResetQuery();
            }
            catch (OutOfMemoryException)
            {
                ShowMessage("The selected query is too large to run.");
                ResetQuery();
            }
            FinishTask();
        }

        private bool SearchCommandCanExecute()
        {
            return !SearchText.IsNullOrWhiteSpace() && SearchTypeIsChecked();
        }

        private async void SearchCommandExecute()
        {
            await RunQuery(GetSearchQueryType(), searchText: SearchText);
        }

        private bool SearchOuCommandCanExecute()
        {
            return OuCommandCanExecute() && SearchCommandCanExecute();
        }

        private async void SearchOuCommandExecute()
        {
            await RunQuery(
                GetSearchQueryType(), CurrentScope, searchText: SearchText);
        }

        private bool SearchTypeIsChecked()
        {
            return ComputerSearchIsChecked ||
                   GroupSearchIsChecked ||
                   UserSearchIsChecked;
        }

        private void SetUpCommands()
        {
            GetOuComputersCommand = new RelayCommand(
                GetOuComputersCommandExecute, OuCommandCanExecute);

            GetOuGroupsCommand = new RelayCommand(
                GetOuGroupsCommandExecute, OuCommandCanExecute);

            GetOuUsersCommand = new RelayCommand(
                GetOuUsersCommandExecute, OuCommandCanExecute);

            GetOuUsersDirectReportsCommand = new RelayCommand(
                GetOuUsersDirectReportsCommandExecute, OuCommandCanExecute);

            GetOuUsersGroupsCommand = new RelayCommand(
                GetOuUsersGroupsCommandExecute, OuCommandCanExecute);

            WriteToFileCommand = new RelayCommand(
                WriteToFileCommandExecute, WriteToFileCommandCanExecute);

            OpenAboutWindowCommand = new RelayCommand(
                OpenAboutWindowCommandExecute);

            OpenHelpWindowCommand = new RelayCommand(
                OpenHelpWindowCommandExecute);

            GetUserGroupsCommand = new RelayCommand(
                GetContextUserGroupsCommandExecute);

            GetComputerGroupsCommand = new RelayCommand(
                GetContextComputerGroupsCommandExecute);

            GetDirectReportDirectReportsCommand = new RelayCommand(
                GetContextDirectReportDirectReportsCommandExecute);

            GetDirectReportGroupsCommand = new RelayCommand(
                GetContextDirectReportGroupsCommandExecute);

            GetGroupComputersCommand = new RelayCommand(
                GetContextGroupComputersCommandExecute);

            GetGroupUsersCommand = new RelayCommand(
                GetContextGroupUsersCommandExecute);

            GetGroupUsersDirectReportsCommand = new RelayCommand(
                GetContextGroupUsersDirectReportsCommandExecute);

            GetUserDirectReportsCommand = new RelayCommand(
                GetContextUserDirectReportsCommandExecute);

            GetGroupUsersGroupsCommand = new RelayCommand(
                GetContextGroupUsersGroupsCommandExecute);

            GetDirectReportSummaryCommand = new RelayCommand(
                GetContextDirectReportSummaryCommandExecute);

            GetUserSummaryCommand = new RelayCommand(
                GetContextUserSummaryCommandExecute);

            GetComputerSummaryCommand = new RelayCommand(
                GetContextComputerSummaryCommandExecute);

            GetGroupSummaryCommand = new RelayCommand(
                GetContextGroupSummaryCommandExecute);

            CancelCommand = new RelayCommand(CancelCommandExecute);

            PreviousQueryCommand = new RelayCommand(
                PreviousQueryCommandExecute, PreviousQueryCommandCanExecute);

            SelectionChangedCommand =
                new RelayCommand<IList>(
                    items =>
                    {
                        SelectedDataRowViews = items;
                    });

            GetOuGroupsUsersCommand = new RelayCommand(
                GetOuGroupsUsersCommandExecute, OuCommandCanExecute);

            SearchCommand = new RelayCommand(
                SearchCommandExecute, SearchCommandCanExecute);

            SearchOuCommand = new RelayCommand(
                SearchOuCommandExecute, SearchOuCommandCanExecute);

            GetManagerDirectReportsCommand = new RelayCommand(
                GetManagerDirectReportsCommandExecute);

            GetManagerGroupsCommand = new RelayCommand(
                GetManagerGroupsCommandExecute);

            GetManagerSummaryCommand = new RelayCommand(
                GetManagerSummaryCommandExecute);

            GetContainerGroupComputersCommand = new RelayCommand(
                GetContainerGroupComputersCommandExecute);

            GetContainerGroupManagedByDirectReportsCommand = new RelayCommand(
                GetContainerGroupManagedByDirectReportsCommandExecute);

            GetContainerGroupSummaryCommand = new RelayCommand(
                GetContainerGroupSummaryCommandExecute);

            GetContainerGroupUsersCommand = new RelayCommand(
                GetContainerGroupUsersCommandExecute);

            GetContainerGroupManagedByGroupsCommand = new RelayCommand(
                GetContainerGroupManagedByGroupsCommandExecute);

            GetContainerGroupManagedBySummaryCommand = new RelayCommand(
                GetContainerGroupManagedBySummaryCommandExecute);

            GetContainerGroupUsersDirectReportsCommand = new RelayCommand(
                GetContainerGroupUsersDirectReportsCommandExecute);
        }

        private void SetViewVariables()
        {
            ProgressBarVisibility = Visibility.Hidden;
            MessageVisibility = Visibility.Hidden;
            CancelButtonVisibility = Visibility.Hidden;
            ViewIsEnabled = true;
            ContextMenuItems = null;
            try
            {
                Version = CurrentDeployment.CurrentVersion.ToString();
            }
            catch (InvalidDeploymentException)
            {
                Version = "Not Installed";
            }
            SetUpCommands();
        }

        private void ShowMessage(string messageContent)
        {
            MessageContent = messageContent + "\n\nDouble-click to dismiss.";
            MessageVisibility = Visibility.Visible;
        }

        private void StartTask()
        {
            HideMessage();
            ViewIsEnabled = false;
            ProgressBarVisibility = Visibility.Visible;
            CancelButtonVisibility = Visibility.Visible;
        }

        private bool WriteToFileCommandCanExecute()
        {
            return Data != null && Data.Count > 0;
        }

        private async void WriteToFileCommandExecute()
        {
            StartTask();
            await Task.Run(
                () =>
                {
                    var fileWriter = new DataFileWriter
                    {
                        Data = Data,
                        Scope = Queries.First().Scope,
                        QueryType = Queries.First().QueryType
                    };
                    ShowMessage("Wrote data to:\n" + fileWriter.WriteToCsv());
                });
            FinishTask();
        }
    }
}