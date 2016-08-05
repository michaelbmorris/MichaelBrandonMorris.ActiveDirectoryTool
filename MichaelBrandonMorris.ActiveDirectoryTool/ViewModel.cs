using System;
using System.Collections;
using System.Collections.Generic;
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

namespace MichaelBrandonMorris.ActiveDirectoryTool
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

        private const string ManagerDistinguishedName =
            "ManagerDistinguishedName";

        private const string UserDistinguishedName = "UserDistinguishedName";

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
        private Stack<Query> _queries;
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

            Queries = new Stack<Query>();
        }

        public Visibility CancelButtonVisibility
        {
            get
            {
                return _cancelButtonVisibility;
            }
            set
            {
                if (_cancelButtonVisibility == value)
                {
                    return;
                }
                _cancelButtonVisibility = value;
                NotifyPropertyChanged();
            }
        }

        public ICommand CancelCommand => new RelayCommand(ExecuteCancel);

        public bool ComputerSearchIsChecked
        {
            get
            {
                return _computerSearchIsChecked;
            }
            set
            {
                if (_computerSearchIsChecked == value)
                {
                    return;
                }
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
                if (_contextMenuItems == value)
                {
                    return;
                }

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
                if (_contextMenuVisibility == value)
                {
                    return;
                }

                _contextMenuVisibility = value;
                NotifyPropertyChanged();
            }
        }

        public Scope CurrentScope
        {
            get;
            set;
        }

        public DataView Data
        {
            get
            {
                return _data;
            }
            private set
            {
                if (_data == value)
                {
                    return;
                }

                _data = value;
                NotifyPropertyChanged();
            }
        }

        public ICommand GetOuComputersCommand => new RelayCommand(
            ExecuteGetOuComputers, CanExecuteOuCommand);

        public ICommand GetOuGroupsCommand => new RelayCommand(
            ExecuteGetOuGroups, CanExecuteOuCommand);

        public ICommand GetOuGroupsUsersCommand => new RelayCommand(
            ExecuteGetOuGroupsUsers, CanExecuteOuCommand);

        public ICommand GetOuUsersCommand => new RelayCommand(
            ExecuteGetOuUsers, CanExecuteOuCommand);

        public ICommand GetOuUsersDirectReportsCommand => new RelayCommand(
            ExecuteGetOuUsersDirectReports, CanExecuteOuCommand);

        public ICommand GetOuUsersGroupsCommand => new RelayCommand(
            ExecuteGetOuUsersGroups, CanExecuteOuCommand);

        public bool GroupSearchIsChecked
        {
            get
            {
                return _groupSearchIsChecked;
            }
            set
            {
                if (_groupSearchIsChecked == value)
                {
                    return;
                }

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
                if (_messageContent == value)
                {
                    return;
                }

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
                if (_messageVisibility == value)
                {
                    return;
                }

                _messageVisibility = value;
                NotifyPropertyChanged();
            }
        }

        public ICommand OpenAboutWindow => new RelayCommand(
            ExecuteOpenAboutWindow);

        public ICommand OpenHelpWindow => new RelayCommand(
            ExecuteOpenHelpWindow);

        public Visibility ProgressBarVisibility
        {
            get
            {
                return _progressBarVisibility;
            }
            set
            {
                if (_progressBarVisibility == value)
                {
                    return;
                }

                _progressBarVisibility = value;
                NotifyPropertyChanged();
            }
        }

        public Stack<Query> Queries
        {
            get
            {
                return _queries;
            }
            set
            {
                if (_queries == value)
                {
                    return;
                }

                _queries = value;
                NotifyPropertyChanged();
            }
        }

        public Scope RootScope
        {
            get;
        }

        public ICommand RunPreviousQuery => new RelayCommand(
            ExecuteRunPreviousQuery, CanExecuteRunPreviousQuery);

        public ICommand Search => new RelayCommand(
            ExecuteSearch, CanExecuteSearch);

        public ICommand SearchOu => new RelayCommand(
            ExecuteSearchOu, CanExecuteSearchOu);

        public string SearchText
        {
            get
            {
                return _searchText;
            }
            set
            {
                if (_searchText == value)
                {
                    return;
                }

                _searchText = value;
                NotifyPropertyChanged();
            }
        }

        public ICommand SelectionChangedCommand => new RelayCommand<IList>(
            items =>
            {
                SelectedDataRowViews = items;
            });

        public bool ShowDistinguishedNames
        {
            get
            {
                return _showDistinguishedNames;
            }
            set
            {
                if (_showDistinguishedNames == value)
                {
                    return;
                }

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
                if (_userSearchIsChecked == value)
                {
                    return;
                }

                _userSearchIsChecked = value;
                NotifyPropertyChanged();
            }
        }

        public string Version
        {
            get;
            private set;
        }

        public bool ViewIsEnabled
        {
            get
            {
                return _viewIsEnabled;
            }
            private set
            {
                if (_viewIsEnabled == value)
                {
                    return;
                }

                _viewIsEnabled = value;
                NotifyPropertyChanged();
            }
        }

        public ICommand WriteToFileCommand => new RelayCommand(
            ExecuteWriteToFile, CanExecuteWriteToFile);

        private AboutWindow AboutWindow
        {
            get
            {
                if (_aboutWindow == null || !_aboutWindow.IsVisible)
                {
                    _aboutWindow = new AboutWindow();
                }

                return _aboutWindow;
            }
        }

        private MenuItem GetContainerGroupsComputers => new MenuItem
        {
            Header = "Container Groups - Get Computers",
            Command = new RelayCommand(ExecuteGetContainerGroupsComputers)
        };

        private MenuItem GetContainerGroupsManagedByDirectReports =>
            new MenuItem
            {
                Header = "Container Group - Get Managed By's Direct Reports",
                Command = new RelayCommand(
                    ExecuteGetContainerGroupsManagedByDirectReports)
            };

        private MenuItem GetContainerGroupsManagedByGroups => new MenuItem
        {
            Header = "Container Groups - Get Managed By's Groups",
            Command = new RelayCommand(
                ExecuteGetContainerGroupsManagedByGroups)
        };

        private MenuItem GetContainerGroupsManagedBySummaries => new MenuItem
        {
            Header = "Container Groups - Get Managed By's Summaries",
            Command = new RelayCommand(
                ExecuteGetContainerGroupsManagedBySummaries)
        };

        private MenuItem GetContainerGroupsSummaries => new MenuItem
        {
            Header = "Container Groups - Get Summaries",
            Command = new RelayCommand(ExecuteGetContainerGroupsSummaries)
        };

        private MenuItem GetContainerGroupsUsers => new MenuItem
        {
            Header = "Container Groups - Get Users",
            Command = new RelayCommand(ExecuteGetContainerGroupsUsers)
        };

        private MenuItem GetContainerGroupsUsersDirectReports => new MenuItem
        {
            Header = "Container Groups - Get Users' Direct Reports",
            Command = new RelayCommand(
                ExecuteGetContainerGroupsUsersDirectReports)
        };

        private MenuItem GetDirectReportsDirectReports => new MenuItem
        {
            Header = "Direct Reports - Get Direct Reports",
            Command = new RelayCommand(ExecuteGetDirectReportsDirectReports)
        };

        private MenuItem GetDirectReportsGroups => new MenuItem
        {
            Header = "Direct Reports - Get Groups",
            Command = new RelayCommand(ExecuteGetDirectReportsGroups)
        };

        private MenuItem GetDirectReportsSummaries => new MenuItem
        {
            Header = "Direct Reports - Get Summaries",
            Command = new RelayCommand(ExecuteGetDirectReportsSummaries)
        };

        private MenuItem GetGroupsComputers => new MenuItem
        {
            Header = "Groups - Get Computers",
            Command = new RelayCommand(ExecuteGetGroupsComputers)
        };

        private MenuItem GetGroupsSummaries => new MenuItem
        {
            Header = "Groups - Get Summaries",
            Command = new RelayCommand(ExecuteGetGroupsSummaries)
        };

        private MenuItem GetGroupsUsers => new MenuItem
        {
            Header = "Groups - Get Users",
            Command = new RelayCommand(ExecuteGetGroupsUsers)
        };

        private MenuItem GetGroupsUsersDirectReports => new MenuItem
        {
            Header = "Groups - Get Users' Direct Reports",
            Command = new RelayCommand(ExecuteGetGroupsUsersDirectReports)
        };

        private MenuItem GetGroupsUsersGroups => new MenuItem
        {
            Header = "Groups - Get Users' Groups",
            Command = new RelayCommand(ExecuteGetGroupsUsersGroups)
        };

        private MenuItem GetManagersDirectReports => new MenuItem
        {
            Header = "Managers - Get Direct Reports",
            Command = new RelayCommand(ExecuteGetManagersDirectReports)
        };

        private MenuItem GetManagersGroups => new MenuItem
        {
            Header = "Managers - Get Groups",
            Command = new RelayCommand(ExecuteGetManagersGroups)
        };

        private MenuItem GetManagersSummaries => new MenuItem
        {
            Header = "Managers - Get Summaries",
            Command = new RelayCommand(ExecuteGetManagersSummaries)
        };

        private MenuItem GetUsersDirectReports => new MenuItem
        {
            Header = "Users - Get Direct Reports",
            Command = new RelayCommand(ExecuteGetUsersDirectReports)
        };

        private MenuItem GetUsersGroups => new MenuItem
        {
            Header = "Users - Get Groups",
            Command = new RelayCommand(ExecuteGetUsersGroups)
        };

        private MenuItem GetUsersSummaries => new MenuItem
        {
            Header = "Users - Get Summaries",
            Command = new RelayCommand(ExecuteGetUsersSummaries)
        };

        private MenuItem MenuItemGetComputersGroups => new MenuItem
        {
            Header = "Computers - Get Groups",
            Command = new RelayCommand(ExecuteGetComputersGroups)
        };

        private MenuItem MenuItemGetComputersSummaries => new MenuItem
        {
            Header = "Computers - Get Summaries",
            Command = new RelayCommand(ExecuteGetComputersSummaries)
        };

        private MenuItem MenuItemGetContainerGroupsUsersGroups => new MenuItem
        {
            Header = "Container Groups - Get Users' Groups",
            Command = new RelayCommand(ExecuteGetContainerGroupsUsersGroups)
        };

        private MenuItem MenuItemGetGroupsManagedByDirectReports
            => new MenuItem
            {
                Header = "Group - Get Managed By's Direct Reports",
                Command = new RelayCommand(
                    ExecuteGetGroupsManagedByDirectReports)
            };

        private MenuItem MenuItemGetGroupsManagedByGroups
            => new MenuItem
            {
                Header = "Group - Get Managed By's Groups",
                Command = new RelayCommand(ExecuteGetGroupsManagedByGroups)
            };

        private MenuItem MenuItemGetGroupsManagedBySummaries
            => new MenuItem
            {
                Header = "Groups - Get Managed By Summaries",
                Command = new RelayCommand(ExecuteGetGroupsManagedBySummaries)
            };

        private IList SelectedDataRowViews
        {
            get;
            set;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private static void ExecuteOpenHelpWindow() => Process.Start(
            HelpFile);

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

        private static string GetGroupManagedByDistinguishedName(
            DataRowView dataRowView)
        {
            return dataRowView[GroupManagedByDistinguishedName].ToString();
        }

        private static string GetManagerDistinguishedName(
            DataRowView dataRowView)
        {
            return dataRowView[ManagerDistinguishedName].ToString();
        }

        private static string GetUserDistinguishedName(DataRowView dataRowView)
        {
            return dataRowView[UserDistinguishedName].ToString();
        }

        private bool CanExecuteOuCommand()
        {
            return CurrentScope != null;
        }

        private bool CanExecuteRunPreviousQuery()
        {
            return Queries.Multiple();
        }

        private bool CanExecuteSearch()
        {
            return !SearchText.IsNullOrWhiteSpace() && SearchTypeIsChecked();
        }

        private bool CanExecuteSearchOu()
        {
            return CanExecuteOuCommand() && CanExecuteSearch();
        }

        private bool CanExecuteWriteToFile()
        {
            return Data != null && Data.Count > 0;
        }

        private void ExecuteCancel()
        {
            Queries.Peek()?.Cancel();
            ResetQuery();
        }

        private async void ExecuteGetComputersGroups()
        {
            await RunQuery(
                QueryType.ComputersGroups,
                CurrentScope,
                GetComputersDistinguishedNames());
        }

        private async void ExecuteGetComputersSummaries()
        {
            await RunQuery(
                QueryType.ComputersSummaries,
                CurrentScope,
                GetComputersDistinguishedNames());
        }

        private async void ExecuteGetContainerGroupsComputers()
        {
            await RunQuery(
                QueryType.GroupsComputers,
                CurrentScope,
                GetContainerGroupsDistinguishedNames());
        }

        private async void ExecuteGetContainerGroupsManagedByDirectReports()
        {
            await RunQuery(
                QueryType.GroupsUsersDirectReports,
                CurrentScope,
                GetContainerGroupsManagedByDistinguishedNames());
        }

        private async void ExecuteGetContainerGroupsManagedByGroups()
        {
            await RunQuery(
                QueryType.UsersGroups,
                CurrentScope,
                GetContainerGroupsManagedByDistinguishedNames());
        }

        private async void ExecuteGetContainerGroupsManagedBySummaries()
        {
            await RunQuery(
                QueryType.UsersSummaries,
                CurrentScope,
                GetContainerGroupsManagedByDistinguishedNames());
        }

        private async void ExecuteGetContainerGroupsSummaries()
        {
            await RunQuery(
                QueryType.GroupsSummaries,
                CurrentScope,
                GetContainerGroupsDistinguishedNames());
        }

        private async void ExecuteGetContainerGroupsUsers()
        {
            await RunQuery(
                QueryType.GroupsUsers,
                CurrentScope,
                GetContainerGroupsDistinguishedNames());
        }

        private async void ExecuteGetContainerGroupsUsersDirectReports()
        {
            await RunQuery(
                QueryType.GroupsUsersDirectReports,
                CurrentScope,
                GetContainerGroupsDistinguishedNames());
        }

        private async void ExecuteGetContainerGroupsUsersGroups()
        {
            await RunQuery(
                QueryType.GroupsUsersGroups,
                CurrentScope,
                GetContainerGroupsDistinguishedNames());
        }

        private async void ExecuteGetDirectReportsDirectReports()
        {
            await RunQuery(
                QueryType.UsersDirectReports,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        private async void ExecuteGetDirectReportsGroups()
        {
            await RunQuery(
                QueryType.UsersGroups,
                CurrentScope,
                GetDirectReportsDistinguishedNames());
        }

        private async void ExecuteGetDirectReportsSummaries()
        {
            await RunQuery(
                QueryType.UsersSummaries,
                CurrentScope,
                GetDirectReportsDistinguishedNames());
        }

        private async void ExecuteGetGroupsComputers()
        {
            await RunQuery(
                QueryType.GroupsComputers,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        private async void ExecuteGetGroupsManagedByDirectReports()
        {
            await RunQuery(
                QueryType.UsersDirectReports,
                CurrentScope,
                GetGroupsManagedByDistinguishedNames());
        }

        private async void ExecuteGetGroupsManagedByGroups()
        {
            await RunQuery(
                QueryType.UsersGroups,
                CurrentScope,
                GetGroupsManagedByDistinguishedNames());
        }

        private async void ExecuteGetGroupsManagedBySummaries()
        {
            await RunQuery(
                QueryType.UsersSummaries,
                CurrentScope,
                GetGroupsManagedByDistinguishedNames());
        }

        private async void ExecuteGetGroupsSummaries()
        {
            await RunQuery(
                QueryType.GroupsSummaries,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        private async void ExecuteGetGroupsUsers()
        {
            await RunQuery(
                QueryType.GroupsUsers,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        private async void ExecuteGetGroupsUsersDirectReports()
        {
            await RunQuery(
                QueryType.GroupsUsersDirectReports,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        private async void ExecuteGetGroupsUsersGroups()
        {
            await RunQuery(
                QueryType.GroupsUsersGroups,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        private async void ExecuteGetManagersDirectReports()
        {
            await RunQuery(
                QueryType.UsersDirectReports,
                CurrentScope,
                GetManagersDistinguishedNames());
        }

        private async void ExecuteGetManagersGroups()
        {
            await RunQuery(
                QueryType.UsersGroups,
                CurrentScope,
                GetManagersDistinguishedNames());
        }

        private async void ExecuteGetManagersSummaries()
        {
            await RunQuery(
                QueryType.UsersSummaries,
                CurrentScope,
                GetManagersDistinguishedNames());
        }

        private async void ExecuteGetOuComputers()
        {
            await RunQuery(QueryType.OuComputers, CurrentScope);
        }

        private async void ExecuteGetOuGroups()
        {
            await RunQuery(QueryType.OuGroups, CurrentScope);
        }

        private async void ExecuteGetOuGroupsUsers()
        {
            await RunQuery(QueryType.OuGroupsUsers, CurrentScope);
        }

        private async void ExecuteGetOuUsers()
        {
            await RunQuery(QueryType.OuUsers, CurrentScope);
        }

        private async void ExecuteGetOuUsersDirectReports()
        {
            await RunQuery(QueryType.OuUsersDirectReports, CurrentScope);
        }

        private async void ExecuteGetOuUsersGroups()
        {
            await RunQuery(QueryType.OuUsersGroups, CurrentScope);
        }

        private async void ExecuteGetUsersDirectReports()
        {
            await RunQuery(
                QueryType.UsersDirectReports,
                CurrentScope,
                GetUsersDistinguishedNames());
        }

        private async void ExecuteGetUsersGroups()
        {
            await RunQuery(
                QueryType.UsersGroups,
                CurrentScope,
                GetUsersDistinguishedNames());
        }

        private async void ExecuteGetUsersSummaries()
        {
            await RunQuery(
                QueryType.UsersSummaries,
                CurrentScope,
                GetUsersDistinguishedNames());
        }

        private void ExecuteOpenAboutWindow()
        {
            AboutWindow.Show();
            AboutWindow.Activate();
        }

        private async void ExecuteRunPreviousQuery()
        {
            Queries.Pop();
            await RunQuery(Queries.Pop());
        }

        private async void ExecuteSearch()
        {
            await RunQuery(GetSearchQueryType(), searchText: SearchText);
        }

        private async void ExecuteSearchOu()
        {
            await RunQuery(
                GetSearchQueryType(), CurrentScope, searchText: SearchText);
        }

        private async void ExecuteWriteToFile()
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
            if (queryType != QueryType.ComputersGroups)
            {
                computerContextMenuItems.Add(MenuItemGetComputersGroups);
            }

            if (queryType != QueryType.ComputersSummaries)
            {
                computerContextMenuItems.Add(MenuItemGetComputersSummaries);
            }

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
                contextMenuItems.Add(GetContainerGroupsComputers);
                contextMenuItems.Add(GetContainerGroupsSummaries);
                contextMenuItems.Add(GetContainerGroupsUsers);
                contextMenuItems.Add(MenuItemGetContainerGroupsUsersGroups);
                contextMenuItems.Add(GetContainerGroupsUsersDirectReports);
            }

            if (Data.Table.Columns.Contains(
                ContainerGroupManagedByDistinguishedName))
            {
                contextMenuItems.Add(GetContainerGroupsManagedByDirectReports);
                contextMenuItems.Add(GetContainerGroupsManagedByGroups);
                contextMenuItems.Add(GetContainerGroupsManagedBySummaries);
            }

            if (Data.Table.Columns.Contains(DirectReportDistinguishedName))
            {
                contextMenuItems.Add(GetDirectReportsDirectReports);
                contextMenuItems.Add(GetDirectReportsGroups);
                contextMenuItems.Add(GetDirectReportsSummaries);
            }

            if (Data.Table.Columns.Contains(GroupDistinguishedName))
            {
                contextMenuItems.AddRange(GenerateGroupContextMenuItems());
            }

            if (Data.Table.Columns.Contains(ManagerDistinguishedName))
            {
                contextMenuItems.Add(GetManagersDirectReports);
                contextMenuItems.Add(GetManagersGroups);
                contextMenuItems.Add(GetManagersSummaries);
            }

            if (!Data.Table.Columns.Contains(UserDistinguishedName))
            {
                return contextMenuItems;
            }

            contextMenuItems.AddRange(GenerateUserContextMenuItems());
            return contextMenuItems;
        }

        private IEnumerable<MenuItem> GenerateGroupContextMenuItems()
        {
            var groupContextMenuItems = new List<MenuItem>();
            var queryType = Queries.Peek().QueryType;
            if (queryType != QueryType.GroupsComputers)
            {
                groupContextMenuItems.Add(GetGroupsComputers);
            }

            if (queryType != QueryType.GroupsSummaries)
            {
                groupContextMenuItems.Add(GetGroupsSummaries);
            }

            if (queryType != QueryType.GroupsUsers)
            {
                groupContextMenuItems.Add(GetGroupsUsers);
            }

            if (queryType != QueryType.GroupsUsersDirectReports)
            {
                groupContextMenuItems.Add(GetGroupsUsersDirectReports);
            }

            if (queryType != QueryType.GroupsUsersGroups)
            {
                groupContextMenuItems.Add(GetGroupsUsersGroups);
            }

            if (Data.Table.Columns.Contains(GroupManagedByDistinguishedName))
            {
                groupContextMenuItems.AddRange(
                    GenerateGroupManagedByContextMenuItems());
            }

            return groupContextMenuItems;
        }

        private IEnumerable<MenuItem> GenerateGroupManagedByContextMenuItems()
        {
            return new List<MenuItem>
            {
                MenuItemGetGroupsManagedByDirectReports,
                MenuItemGetGroupsManagedByGroups,
                MenuItemGetGroupsManagedBySummaries
            };
        }

        private IEnumerable<MenuItem> GenerateUserContextMenuItems()
        {
            var userContextMenuItems = new List<MenuItem>();
            var queryType = Queries.Peek().QueryType;
            if (queryType != QueryType.UsersDirectReports)
            {
                userContextMenuItems.Add(GetUsersDirectReports);
            }

            if (queryType != QueryType.UsersGroups)
            {
                userContextMenuItems.Add(GetUsersGroups);
            }

            if (queryType != QueryType.UsersSummaries)
            {
                userContextMenuItems.Add(GetUsersSummaries);
            }

            return userContextMenuItems;
        }

        private IEnumerable<string> GetComputersDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>().Select(
                GetComputerDistinguishedName).Where(
                    c => !c.IsNullOrWhiteSpace());
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

        private IEnumerable<string> GetGroupsManagedByDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>().Select(
                GetGroupManagedByDistinguishedName).Where(
                    s => !s.IsNullOrWhiteSpace());
        }

        private IEnumerable<string> GetManagersDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>().Select(
                GetManagerDistinguishedName).Where(
                    m => !m.IsNullOrWhiteSpace());
        }

        private QueryType GetSearchQueryType()
        {
            if (ComputerSearchIsChecked)
            {
                return QueryType.SearchComputer;
            }

            if (GroupSearchIsChecked)
            {
                return QueryType.SearchGroup;
            }

            return UserSearchIsChecked ? QueryType.SearchUser : QueryType.None;
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

        private void ResetQuery()
        {
            if (Queries.Any())
            {
                Queries.Pop();
            }
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
                    selectedItemDistinguishedNames?.ToList(),
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

        private bool SearchTypeIsChecked()
        {
            return ComputerSearchIsChecked ||
                   GroupSearchIsChecked ||
                   UserSearchIsChecked;
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
                Version = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            catch (InvalidDeploymentException)
            {
                Version = "Not Installed";
            }
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
    }
}