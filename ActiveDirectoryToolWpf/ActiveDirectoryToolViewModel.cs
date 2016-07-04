using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Deployment.Application;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using GalaSoft.MvvmLight.CommandWpf;
using static System.Deployment.Application.ApplicationDeployment;

namespace ActiveDirectoryToolWpf
{
    public enum QueryType
    {
        ContextualComputerGroups,
        ContextualDirectReportDirectReports,
        ContextualDirectReportGroups,
        ContextualGroupComputers,
        ContextualGroupUsersDirectReports,
        ContextualGroupUsers,
        ContextualUserDirectReports,
        ContextualUserGroups,
        OrganizationalUnitComputers,
        OrganizationalUnitGroups,
        OrganizationalUnitUsers,
        OrganizationalUnitUsersDirectReports,
        OrganizationalUnitUsersGroups
    }

    public class ActiveDirectoryToolViewModel : INotifyPropertyChanged
    {
        private static readonly ActiveDirectoryAttribute[]
            DefaultComputerAttributes =
            {
                ActiveDirectoryAttribute.ComputerName,
                ActiveDirectoryAttribute.ComputerDescription,
                ActiveDirectoryAttribute.ComputerDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultComputerGroupsAttributes =
            {
                ActiveDirectoryAttribute.ComputerName,
                ActiveDirectoryAttribute.GroupName,
                ActiveDirectoryAttribute.ComputerDistinguishedName,
                ActiveDirectoryAttribute.GroupDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultGroupAttributes =
            {
                ActiveDirectoryAttribute.GroupSamAccountName,
                ActiveDirectoryAttribute.GroupManagedBy,
                ActiveDirectoryAttribute.GroupDescription,
                ActiveDirectoryAttribute.GroupDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultGroupComputersAttributes =
            {
                ActiveDirectoryAttribute.GroupName,
                ActiveDirectoryAttribute.ComputerName,
                ActiveDirectoryAttribute.GroupDistinguishedName,
                ActiveDirectoryAttribute.ComputerDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultGroupUsersAttributes =
            {
                ActiveDirectoryAttribute.GroupName,
                ActiveDirectoryAttribute.UserSurname,
                ActiveDirectoryAttribute.UserGivenName,
                ActiveDirectoryAttribute.UserDisplayName,
                ActiveDirectoryAttribute.UserSamAccountName,
                ActiveDirectoryAttribute.UserIsActive,
                ActiveDirectoryAttribute.UserIsAccountLockedOut,
                ActiveDirectoryAttribute.UserDescription,
                ActiveDirectoryAttribute.UserTitle,
                ActiveDirectoryAttribute.UserCompany,
                ActiveDirectoryAttribute.UserManager,
                ActiveDirectoryAttribute.UserHomeDrive,
                ActiveDirectoryAttribute.UserHomeDirectory,
                ActiveDirectoryAttribute.UserScriptPath,
                ActiveDirectoryAttribute.UserEmailAddress,
                ActiveDirectoryAttribute.UserStreetAddress,
                ActiveDirectoryAttribute.UserCity,
                ActiveDirectoryAttribute.UserState,
                ActiveDirectoryAttribute.UserVoiceTelephoneNumber,
                ActiveDirectoryAttribute.UserPager,
                ActiveDirectoryAttribute.UserMobile,
                ActiveDirectoryAttribute.UserFax,
                ActiveDirectoryAttribute.UserVoip,
                ActiveDirectoryAttribute.UserSip,
                ActiveDirectoryAttribute.UserUserPrincipalName,
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.GroupDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultGroupUsersDirectReportsAttributes =
            {
                ActiveDirectoryAttribute.GroupName,
                ActiveDirectoryAttribute.UserDisplayName,
                ActiveDirectoryAttribute.UserSamAccountName,
                ActiveDirectoryAttribute.DirectReportDisplayName,
                ActiveDirectoryAttribute.DirectReportSamAccountName,
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.DirectReportDistinguishedName,
                ActiveDirectoryAttribute.GroupDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultUserAttributes =
            {
                ActiveDirectoryAttribute.UserSurname,
                ActiveDirectoryAttribute.UserGivenName,
                ActiveDirectoryAttribute.UserDisplayName,
                ActiveDirectoryAttribute.UserSamAccountName,
                ActiveDirectoryAttribute.UserIsActive,
                ActiveDirectoryAttribute.UserIsAccountLockedOut,
                ActiveDirectoryAttribute.UserDescription,
                ActiveDirectoryAttribute.UserTitle,
                ActiveDirectoryAttribute.UserCompany,
                ActiveDirectoryAttribute.UserManager,
                ActiveDirectoryAttribute.UserHomeDrive,
                ActiveDirectoryAttribute.UserHomeDirectory,
                ActiveDirectoryAttribute.UserScriptPath,
                ActiveDirectoryAttribute.UserEmailAddress,
                ActiveDirectoryAttribute.UserStreetAddress,
                ActiveDirectoryAttribute.UserCity,
                ActiveDirectoryAttribute.UserState,
                ActiveDirectoryAttribute.UserVoiceTelephoneNumber,
                ActiveDirectoryAttribute.UserPager,
                ActiveDirectoryAttribute.UserMobile,
                ActiveDirectoryAttribute.UserFax,
                ActiveDirectoryAttribute.UserVoip,
                ActiveDirectoryAttribute.UserSip,
                ActiveDirectoryAttribute.UserUserPrincipalName,
                ActiveDirectoryAttribute.UserDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultUserDirectReportsAttributes =
            {
                ActiveDirectoryAttribute.UserDisplayName,
                ActiveDirectoryAttribute.UserSamAccountName,
                ActiveDirectoryAttribute.DirectReportDisplayName,
                ActiveDirectoryAttribute.DirectReportSamAccountName,
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.DirectReportDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultUserGroupsAttributes =
            {
                ActiveDirectoryAttribute.UserSamAccountName,
                ActiveDirectoryAttribute.GroupSamAccountName,
                ActiveDirectoryAttribute.UserName,
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.GroupDistinguishedName
            };

        private readonly MenuItem _computerGetGroupsMenuItem;
        private readonly MenuItem _directReportGetDirectReportsMenuItem;
        private readonly MenuItem _directReportGetGroupsMenuItem;
        private readonly MenuItem _groupGetComputersMenuItem;
        private readonly MenuItem _groupGetUsersDirectReportsMenuItem;
        private readonly MenuItem _groupGetUsersGroupsMenuItem;
        private readonly MenuItem _groupGetUsersMenuItem;

        private readonly PrincipalContext _principalContext;
        private readonly MenuItem _userGetDirectReportsMenuItem;
        private readonly MenuItem _userGetGroupsMenuItem;
        private AboutWindow _aboutWindow;
        private ActiveDirectorySearcher _activeDirectorySearcher;
        private List<MenuItem> _contextMenu;
        private DataView _data;
        private DataPreparer _dataPreparer;
        private HelpWindow _helpWindow;

        private string _messageContent;

        private Visibility _messageVisibility;
        private Visibility _progressBarVisibility;

        private QueryType _queryType;
        private DataRowView _selectedDataGridRow;
        private bool _viewIsEnabled;

        public ActiveDirectoryToolViewModel()
        {
            RootScope = new ActiveDirectoryScopeFetcher().Scope;
            SetViewVariables();
            _principalContext = new PrincipalContext(ContextType.Domain);
            _computerGetGroupsMenuItem = new MenuItem
            {
                Header = "Computer - Get Groups",
                Command = GetComputerGroupsCommand
            };
            _userGetGroupsMenuItem = new MenuItem
            {
                Header = "User - Get Groups",
                Command = GetUserGroupsCommand
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
        }

        public List<MenuItem> ContextMenu
        {
            get { return _contextMenu; }
            private set
            {
                _contextMenu = value;
                NotifyPropertyChanged();
            }
        }

        public ActiveDirectoryScope CurrentScope
        {
            get;
            set;
        }

        public DataView Data
        {
            get { return _data; }
            private set
            {
                _data = value;
                NotifyPropertyChanged();
            }
        }

        public ICommand GetOuComputersCommand
        {
            get;
            private set;
        }

        public ICommand GetOuGroupsCommand
        {
            get;
            private set;
        }

        public ICommand GetOuUsersCommand
        {
            get;
            private set;
        }

        public ICommand GetOuUsersDirectReportsCommand
        {
            get;
            private set;
        }

        public ICommand GetOuUsersGroupsCommand
        {
            get;
            private set;
        }

        public string MessageContent
        {
            get { return _messageContent; }
            set
            {
                _messageContent = value;
                NotifyPropertyChanged();
            }
        }

        public Visibility MessageVisibility
        {
            get { return _messageVisibility; }
            set
            {
                _messageVisibility = value;
                NotifyPropertyChanged();
            }
        }

        public ICommand OpenAboutWindowCommand
        {
            get;
            private set;
        }

        public ICommand OpenHelpWindowCommand
        {
            get;
            private set;
        }

        public Visibility ProgressBarVisibility
        {
            get { return _progressBarVisibility; }
            set
            {
                _progressBarVisibility = value;
                NotifyPropertyChanged();
            }
        }

        public QueryType QueryType
        {
            get { return _queryType; }
            set
            {
                _queryType = value;
                NotifyPropertyChanged();
            }
        }

        public ActiveDirectoryScope RootScope
        {
            get;
        }

        public DataRowView SelectedDataGridRow
        {
            private get { return _selectedDataGridRow; }
            set
            {
                _selectedDataGridRow = value;
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
            get { return _viewIsEnabled; }
            private set
            {
                _viewIsEnabled = value;
                NotifyPropertyChanged();
            }
        }

        public ICommand WriteToFileCommand
        {
            get;
            private set;
        }

        private ICommand GetComputerGroupsCommand
        {
            get;
            set;
        }

        private ICommand GetDirectReportDirectReportsCommand
        {
            get;
            set;
        }

        private ICommand GetDirectReportGroupsCommand
        {
            get;
            set;
        }

        private ICommand GetGroupComputersCommand
        {
            get;
            set;
        }

        private ICommand GetGroupUsersCommand
        {
            get;
            set;
        }

        private ICommand GetGroupUsersDirectReportsCommand
        {
            get;
            set;
        }

        private ICommand GetGroupUsersGroupsCommand
        {
            get;
            set;
        }

        private ICommand GetUserGroupsCommand
        {
            get;
            set;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void FinishTask()
        {
            ContextMenu = GenerateContextMenu();
            ProgressBarVisibility = Visibility.Hidden;
            ViewIsEnabled = true;
        }

        private List<MenuItem> GenerateContextMenu()
        {
            var contextMenu = new List<MenuItem>();
            if (QueryType == QueryType.OrganizationalUnitComputers)
            {
                contextMenu.Add(_computerGetGroupsMenuItem);
            }
            if (QueryType == QueryType.OrganizationalUnitGroups ||
                QueryType == QueryType.OrganizationalUnitUsersGroups)
            {
                contextMenu.Add(_groupGetComputersMenuItem);
                contextMenu.Add(_groupGetUsersDirectReportsMenuItem);
            }
            if (QueryType == QueryType.OrganizationalUnitUsers ||
                QueryType == QueryType.OrganizationalUnitUsersDirectReports)
            {
                contextMenu.Add(_userGetGroupsMenuItem);
            }
            if (QueryType == QueryType.OrganizationalUnitUsersDirectReports)
            {
                contextMenu.Add(_directReportGetDirectReportsMenuItem);
                contextMenu.Add(_directReportGetGroupsMenuItem);
            }
            if (QueryType == QueryType.OrganizationalUnitUsersGroups)
            {
            }

            return contextMenu;
        }

        private async void GetComputerGroupsCommandExecute()
        {
            StartTask();
            QueryType = QueryType.ContextualComputerGroups;
            await Task.Run(() =>
            {
                var computerPrincipal = ComputerPrincipal.FindByIdentity(
                    _principalContext, GetSelectedComputerDistinguishedName());
                _dataPreparer = new DataPreparer
                {
                    Data = new[]
                    {
                        ActiveDirectorySearcher.GetComputerGroups(
                            computerPrincipal)
                    },
                    Attributes = DefaultComputerGroupsAttributes.ToList()
                };
                try
                {
                    Data = new List<ExpandoObject>(_dataPreparer.GetResults())
                        .ToDataTable().AsDataView();
                }
                catch (ArgumentNullException)
                {
                    ShowMessage(
                        "No groups found for the selected computer.");
                }
            });
            FinishTask();
        }

        private async void GetDirectReportDirectReportsCommandExecute()
        {
            StartTask();
            QueryType = QueryType.ContextualDirectReportDirectReports;
            await Task.Run(() =>
            {
                var directReportUserPrincipal = UserPrincipal.FindByIdentity(
                    _principalContext,
                    GetSelectedDirectReportDistinguishedName());
                _dataPreparer = new DataPreparer
                {
                    Data = new[]
                    {
                        ActiveDirectorySearcher.GetUserDirectReports(
                            directReportUserPrincipal)
                    },
                    Attributes = DefaultUserDirectReportsAttributes.ToList()
                };
                try
                {
                    Data = new List<ExpandoObject>(_dataPreparer.GetResults())
                        .ToDataTable().AsDataView();
                }
                catch (ArgumentNullException)
                {
                    ShowMessage(
                        "No direct reports found for the selected user.");
                }
            });
            FinishTask();
        }

        private async void GetDirectReportGroupsCommandExecute()
        {
            StartTask();
            QueryType = QueryType.ContextualDirectReportDirectReports;
            await Task.Run(() =>
            {
                var directReportUserPrincipal = UserPrincipal.FindByIdentity(
                    _principalContext,
                    GetSelectedDirectReportDistinguishedName());
                _dataPreparer = new DataPreparer
                {
                    Data = new[]
                    {
                        ActiveDirectorySearcher.GetUserGroups(
                            directReportUserPrincipal)
                    },
                    Attributes = DefaultUserGroupsAttributes.ToList()
                };
                try
                {
                    Data = new List<ExpandoObject>(_dataPreparer.GetResults())
                        .ToDataTable().AsDataView();
                }
                catch (ArgumentNullException)
                {
                    ShowMessage(
                        "No groups found for the selected direct report.");
                }
            });
            FinishTask();
        }

        private async void GetGroupComputersCommandExecute()
        {
            StartTask();
            QueryType = QueryType.ContextualGroupComputers;
            await Task.Run(() =>
            {
                var groupPrincipal = GroupPrincipal.FindByIdentity(
                    _principalContext,
                    GetSelectedGroupDistinguishedName());
                _dataPreparer = new DataPreparer
                {
                    Data = new[]
                    {
                        ActiveDirectorySearcher.GetComputersFromGroup(
                            groupPrincipal)
                    },
                    Attributes = DefaultGroupComputersAttributes.ToList()
                };
                try
                {
                    Data = new List<ExpandoObject>(_dataPreparer.GetResults())
                        .ToDataTable().AsDataView();
                }
                catch (ArgumentNullException)
                {
                    ShowMessage(
                        "No groups found for the selected direct report.");
                }
            });
            FinishTask();
        }

        private async void GetGroupUsersCommandExecute()
        {
        }

        private async void GetGroupUsersDirectReportsCommandExecute()
        {
            StartTask();
            QueryType = QueryType.ContextualGroupUsersDirectReports;
            await Task.Run(() =>
            {
                var groupPrincipal = GroupPrincipal.FindByIdentity(
                    _principalContext,
                    GetSelectedGroupDistinguishedName());
                _dataPreparer = new DataPreparer
                {
                    Data = new[]
                    {
                        ActiveDirectorySearcher.GetUsersDirectReports(
                            groupPrincipal)
                    },
                    Attributes = DefaultGroupUsersDirectReportsAttributes
                        .ToList()
                };
                try
                {
                    Data = new List<ExpandoObject>(_dataPreparer.GetResults())
                        .ToDataTable().AsDataView();
                }
                catch (ArgumentNullException)
                {
                    ShowMessage(
                        "No users and/or direct reports found in the selected group.");
                }
            });
            FinishTask();
        }

        private async void GetGroupUsersGroupsCommandExecute()
        {
            /*StartTask();
            QueryType = QueryType.ContextualGroupUsersDirectReports;
            await Task.Run(() =>
            {
                var groupPrincipal = GroupPrincipal.FindByIdentity(
                    _principalContext,
                    GetSelectedGroupDistinguishedName());
                _dataPreparer = new DataPreparer
                {
                    Data = new[]
                    {
                        ActiveDirectorySearcher.GetUsersGroups(
                            groupPrincipal)
                    },
                    Attributes = DefaultGroupUsersGroupsAttributes
                    .ToList()
                };
                try
                {
                    Data = new List<ExpandoObject>(_dataPreparer.GetResults())
                        .ToDataTable().AsDataView();
                }
                catch (ArgumentNullException)
                {
                    ShowMessage(
                        "No users and/or groups found in the selected group.");
                }
            });
            FinishTask();*/
        }

        private async void GetOuComputersCommandExecute()
        {
            StartOuTask();
            QueryType = QueryType.OrganizationalUnitComputers;
            await Task.Run(() =>
            {
                _dataPreparer = new DataPreparer
                {
                    Data = _activeDirectorySearcher.GetOuComputers(),
                    Attributes = DefaultComputerAttributes.ToList()
                };
                try
                {
                    Data = new List<ExpandoObject>(_dataPreparer.GetResults())
                        .ToDataTable().AsDataView();
                }
                catch (ArgumentNullException)
                {
                    ShowMessage("No computers found in selected OU.");
                }
            });
            FinishTask();
        }

        private async void GetOuGroupsCommandExecute()
        {
            StartOuTask();
            QueryType = QueryType.OrganizationalUnitGroups;
            await Task.Run(() =>
            {
                _dataPreparer = new DataPreparer
                {
                    Data = _activeDirectorySearcher.GetOuGroups(),
                    Attributes = DefaultGroupAttributes.ToList()
                };
                try
                {
                    Data = new List<ExpandoObject>(_dataPreparer.GetResults())
                        .ToDataTable().AsDataView();
                }
                catch (ArgumentNullException)
                {
                    ShowMessage("No groups found in selected OU.");
                }
            });
            FinishTask();
        }

        private async void GetOuUsersCommandExecute()
        {
            StartOuTask();
            QueryType = QueryType.OrganizationalUnitUsers;
            await Task.Run(() =>
            {
                _dataPreparer = new DataPreparer
                {
                    Data = _activeDirectorySearcher.GetOuUsers(),
                    Attributes = DefaultUserAttributes.ToList()
                };
                try
                {
                    Data = new List<ExpandoObject>(_dataPreparer.GetResults())
                        .ToDataTable().AsDataView();
                }
                catch (ArgumentNullException)
                {
                    ShowMessage("No users found in selected OU.");
                }
            });
            FinishTask();
        }

        private async void GetOuUsersDirectReportsCommandExecute()
        {
            StartOuTask();
            QueryType = QueryType.OrganizationalUnitUsersDirectReports;
            await Task.Run(() =>
            {
                _dataPreparer = new DataPreparer
                {
                    Data = _activeDirectorySearcher.GetOuUsersDirectReports(),
                    Attributes = DefaultUserDirectReportsAttributes.ToList()
                };
                try
                {
                    Data = new List<ExpandoObject>(_dataPreparer.GetResults())
                        .ToDataTable().AsDataView();
                }
                catch (ArgumentNullException)
                {
                    ShowMessage(
                        "No users and/or direct reports found in selected OU.");
                }
            });
            FinishTask();
        }

        private async void GetOuUsersGroupsCommandExecute()
        {
            StartOuTask();
            QueryType = QueryType.OrganizationalUnitUsersGroups;
            await Task.Run(() =>
            {
                _dataPreparer = new DataPreparer
                {
                    Data = _activeDirectorySearcher.GetOuUsersGroups(),
                    Attributes = DefaultUserGroupsAttributes.ToList()
                };
                try
                {
                    Data = new List<ExpandoObject>(_dataPreparer.GetResults())
                        .ToDataTable().AsDataView();
                }
                catch (ArgumentNullException)
                {
                    ShowMessage(
                        "No users and/or groups found in selected OU.");
                }
            });
            FinishTask();
        }

        private string GetSelectedComputerDistinguishedName()
        {
            return SelectedDataGridRow["ComputerDistinguishedName"].ToString();
        }

        private string GetSelectedDirectReportDistinguishedName()
        {
            return SelectedDataGridRow["DirectReportDistinguishedName"]
                .ToString();
        }

        private string GetSelectedGroupDistinguishedName()
        {
            return SelectedDataGridRow["GroupDistinguishedName"].ToString();
        }

        private string GetSelectedUserDistinguishedName()
        {
            return SelectedDataGridRow["UserDistinguishedName"].ToString();
        }

        private async void GetUserGroupsCommandExecute()
        {
            StartTask();
            QueryType = QueryType.ContextualUserGroups;
            await Task.Run(() =>
            {
                var userPrincipal = UserPrincipal.FindByIdentity(
                    _principalContext, GetSelectedUserDistinguishedName());
                _dataPreparer = new DataPreparer
                {
                    Data = new[]
                    {
                        ActiveDirectorySearcher.GetUserGroups(
                            userPrincipal)
                    },
                    Attributes = DefaultUserGroupsAttributes.ToList()
                };
                try
                {
                    Data = new List<ExpandoObject>(_dataPreparer.GetResults())
                        .ToDataTable().AsDataView();
                }
                catch (ArgumentNullException)
                {
                    ShowMessage(
                        "No groups found for the selected user.");
                }
            });
            FinishTask();
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
                this, new PropertyChangedEventArgs(propertyName));
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

        private void OpenHelpWindowCommandExecute()
        {
            if (_helpWindow == null || !_helpWindow.IsVisible)
            {
                _helpWindow = new HelpWindow();
                _helpWindow.Show();
            }
            else
                _helpWindow.Activate();
        }

        private bool OuCommandCanExecute()
        {
            return CurrentScope != null;
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
                GetUserGroupsCommandExecute);
            GetComputerGroupsCommand = new RelayCommand(
                GetComputerGroupsCommandExecute);
            GetDirectReportDirectReportsCommand = new RelayCommand(
                GetDirectReportDirectReportsCommandExecute);
            GetDirectReportGroupsCommand = new RelayCommand(
                GetDirectReportGroupsCommandExecute);
            GetGroupComputersCommand = new RelayCommand(
                GetGroupComputersCommandExecute);
            GetGroupUsersDirectReportsCommand = new RelayCommand(
                GetGroupUsersDirectReportsCommandExecute);
        }

        private void SetViewVariables()
        {
            ProgressBarVisibility = Visibility.Hidden;
            MessageVisibility = Visibility.Hidden;
            ViewIsEnabled = true;
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

        private void StartOuTask()
        {
            StartTask();
            _activeDirectorySearcher = new ActiveDirectorySearcher(
                CurrentScope);
        }

        private void StartTask()
        {
            HideMessage();
            ViewIsEnabled = false;
            ProgressBarVisibility = Visibility.Visible;
        }

        private bool WriteToFileCommandCanExecute()
        {
            return Data != null && Data.Count > 0;
        }

        private async void WriteToFileCommandExecute()
        {
            StartTask();
            await Task.Run(() =>
            {
                var fileWriter = new DataFileWriter
                {
                    Data = Data,
                    Scope = CurrentScope.Context,
                    QueryType = QueryType
                };
                ShowMessage("Wrote data to:\n" + fileWriter.WriteToCsv());
            });
            FinishTask();
        }
    }
}