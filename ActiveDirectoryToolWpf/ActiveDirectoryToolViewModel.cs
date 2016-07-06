using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Deployment.Application;
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
    public class ActiveDirectoryToolViewModel : INotifyPropertyChanged
    {
        private readonly MenuItem _computerGetGroupsMenuItem;
        private readonly MenuItem _computerGetSummaryMenuItem;
        private readonly MenuItem _directReportGetDirectReportsMenuItem;
        private readonly MenuItem _directReportGetGroupsMenuItem;
        private readonly MenuItem _directReportGetSummaryMenuItem;
        private readonly MenuItem _groupGetComputersMenuItem;
        private readonly MenuItem _groupGetUsersDirectReportsMenuItem;
        private readonly MenuItem _groupGetUsersGroupsMenuItem;
        private readonly MenuItem _groupGetUsersMenuItem;
        private readonly MenuItem _groupGetSummaryMenuItem;
        private readonly MenuItem _userGetDirectReportsMenuItem;
        private readonly MenuItem _userGetGroupsMenuItem;
        private readonly MenuItem _userGetSummaryMenuItem;
        private AboutWindow _aboutWindow;
        private List<MenuItem> _contextMenuItems;
        private DataView _data;
        private HelpWindow _helpWindow;

        private string _messageContent;

        private Visibility _cancelButtonVisibility;
        private Visibility _messageVisibility;
        private Visibility _progressBarVisibility;

        private ObservableStack<ActiveDirectoryQuery> _previousQueries;
        private DataRowView _selectedDataGridRow;
        private bool _viewIsEnabled;

        private ActiveDirectoryQuery _activeDirectoryQuery;

        public ActiveDirectoryQuery Query
        {
            get { return _activeDirectoryQuery;}
            set
            {
                _activeDirectoryQuery = value;
                NotifyPropertyChanged();
            }
        }

        public ObservableStack<ActiveDirectoryQuery> PreviousQueries
        {
            get { return _previousQueries; }
            set
            {
                _previousQueries = value;
                NotifyPropertyChanged();
            }
        }

        public ICommand PreviousQueryCommand
        {
            get;
            private set;
        }

        public ActiveDirectoryToolViewModel()
        {
            RootScope = new ActiveDirectoryScopeFetcher().Scope;
            PreviousQueries = new ObservableStack<ActiveDirectoryQuery>();
            SetViewVariables();
            _computerGetGroupsMenuItem = new MenuItem
            {
                Header = "Computer - Get Groups",
                Command = GetContextComputerGroupsCommand
            };
            _computerGetSummaryMenuItem = new MenuItem
            {
                Header = "Computer - Get Summary",
                Command = GetContextComputerSummaryCommand
            };
            _directReportGetDirectReportsMenuItem = new MenuItem
            {
                Header = "Direct Report - Get Direct Reports",
                Command = GetContextDirectReportDirectReportsCommand
            };
            _directReportGetGroupsMenuItem = new MenuItem
            {
                Header = "Direct Report - Get Groups",
                Command = GetContextDirectReportGroupsCommand
            };
            _directReportGetSummaryMenuItem = new MenuItem
            {
                Header = "Direct Report - Get Summary",
                Command = GetContextDirectReportSummaryCommand
            };
            _groupGetComputersMenuItem = new MenuItem
            {
                Header = "Group - Get Computers",
                Command = GetContextGroupComputersCommand
            };
            _groupGetUsersMenuItem = new MenuItem
            {
                Header = "Group - Get Users",
                Command = GetContextGroupUsersCommand
            };
            _groupGetUsersDirectReportsMenuItem = new MenuItem
            {
                Header = "Group - Get Users' Direct Reports",
                Command = GetContextGroupUsersDirectReportsCommand
            };
            _groupGetUsersGroupsMenuItem = new MenuItem
            {
                Header = "Group - Get Users' Groups",
                Command = GetContextGroupUsersGroupsCommand
            };
            _groupGetSummaryMenuItem = new MenuItem
            {
                Header = "Group - Get Summary",
                Command = GetContextGroupSummaryCommand
            };
            _userGetDirectReportsMenuItem = new MenuItem
            {
                Header = "User - Get Direct Reports",
                Command = GetContextUserDirectReportsCommand
            };
            _userGetGroupsMenuItem = new MenuItem
            {
                Header = "User - Get Groups",
                Command = GetContextUserGroupsCommand
            };
            _userGetSummaryMenuItem = new MenuItem
            {
                Header = "User - Get Summary",
                Command = GetContextUserSummaryCommand
            };
        }

        public Visibility CancelButtonVisibility
        {
            get { return _cancelButtonVisibility;}
            set
            {
                _cancelButtonVisibility = value;
                NotifyPropertyChanged();
            }
        }

        public List<MenuItem> ContextMenuItems
        {
            get { return _contextMenuItems; }
            private set
            {
                _contextMenuItems = value;
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

        public ICommand CancelCommand
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

        private ICommand GetContextComputerGroupsCommand
        {
            get;
            set;
        }

        private ICommand GetContextComputerSummaryCommand
        {
            get;
            set;
        }

        private ICommand GetContextDirectReportDirectReportsCommand
        {
            get;
            set;
        }

        private ICommand GetContextDirectReportGroupsCommand
        {
            get;
            set;
        }

        private ICommand GetContextDirectReportSummaryCommand
        {
            get;
            set;
        }

        private ICommand GetContextGroupComputersCommand
        {
            get;
            set;
        }

        private ICommand GetContextGroupUsersCommand
        {
            get;
            set;
        }

        private ICommand GetContextGroupUsersDirectReportsCommand
        {
            get;
            set;
        }

        private ICommand GetContextGroupUsersGroupsCommand
        {
            get;
            set;
        }

        private ICommand GetContextUserDirectReportsCommand
        {
            get;
            set;
        }

        private ICommand GetContextUserGroupsCommand
        {
            get;
            set;
        }

        private ICommand GetContextUserSummaryCommand
        {
            get;
            set;
        }

        private ICommand GetContextGroupSummaryCommand
        {
            get;
            set;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void FinishTask()
        {
            ContextMenuItems = GenerateContextMenuItems();
            ProgressBarVisibility = Visibility.Hidden;
            CancelButtonVisibility = Visibility.Hidden;
            ViewIsEnabled = true;
        }

        private List<MenuItem> GenerateContextMenuItems()
        {
            var contextMenuItems = new List<MenuItem>();
            switch (_activeDirectoryQuery.QueryType)
            {
                case QueryType.ContextComputerGroups:
                    contextMenuItems.Add(_groupGetComputersMenuItem);
                    contextMenuItems.Add(_groupGetUsersMenuItem);
                    contextMenuItems.Add(_groupGetUsersDirectReportsMenuItem);
                    contextMenuItems.Add(_groupGetUsersGroupsMenuItem);
                    break;
                case QueryType.ContextDirectReportDirectReports:
                    contextMenuItems.Add(
                        _directReportGetDirectReportsMenuItem);
                    contextMenuItems.Add(_directReportGetGroupsMenuItem);
                    contextMenuItems.Add(_directReportGetSummaryMenuItem);
                    contextMenuItems.Add(_userGetGroupsMenuItem);
                    contextMenuItems.Add(_userGetSummaryMenuItem);
                    break;
                case QueryType.ContextDirectReportGroups:
                    contextMenuItems.Add(_groupGetComputersMenuItem);
                    contextMenuItems.Add(_groupGetUsersMenuItem);
                    contextMenuItems.Add(_groupGetUsersDirectReportsMenuItem);
                    contextMenuItems.Add(_groupGetUsersGroupsMenuItem);
                    contextMenuItems.Add(_userGetDirectReportsMenuItem);
                    contextMenuItems.Add(_userGetSummaryMenuItem);
                    break;
                case QueryType.ContextGroupComputers:
                    contextMenuItems.Add(_computerGetGroupsMenuItem);
                    contextMenuItems.Add(_groupGetUsersMenuItem);
                    contextMenuItems.Add(_groupGetUsersDirectReportsMenuItem);
                    contextMenuItems.Add(_groupGetUsersGroupsMenuItem);
                    break;
                case QueryType.ContextGroupUsers:
                    contextMenuItems.Add(_groupGetComputersMenuItem);
                    contextMenuItems.Add(_groupGetUsersDirectReportsMenuItem);
                    contextMenuItems.Add(_groupGetUsersGroupsMenuItem);
                    contextMenuItems.Add(_userGetDirectReportsMenuItem);
                    contextMenuItems.Add(_userGetGroupsMenuItem);
                    contextMenuItems.Add(_userGetSummaryMenuItem);
                    break;
                case QueryType.ContextGroupUsersDirectReports:
                    contextMenuItems.Add(
                        _directReportGetDirectReportsMenuItem);
                    contextMenuItems.Add(_directReportGetGroupsMenuItem);
                    contextMenuItems.Add(_directReportGetSummaryMenuItem);
                    contextMenuItems.Add(_groupGetComputersMenuItem);
                    contextMenuItems.Add(_groupGetUsersMenuItem);
                    contextMenuItems.Add(_groupGetUsersGroupsMenuItem);
                    contextMenuItems.Add(_userGetDirectReportsMenuItem);
                    contextMenuItems.Add(_userGetGroupsMenuItem);
                    contextMenuItems.Add(_userGetSummaryMenuItem);
                    break;
                case QueryType.ContextGroupUsersGroups:
                    contextMenuItems.Add(_groupGetComputersMenuItem);
                    contextMenuItems.Add(_groupGetUsersMenuItem);
                    contextMenuItems.Add(_groupGetUsersGroupsMenuItem);
                    contextMenuItems.Add(_userGetDirectReportsMenuItem);
                    contextMenuItems.Add(_userGetGroupsMenuItem);
                    contextMenuItems.Add(_userGetSummaryMenuItem);
                    break;
                case QueryType.ContextUserDirectReports:
                    contextMenuItems.Add(
                        _directReportGetDirectReportsMenuItem);
                    contextMenuItems.Add(_directReportGetGroupsMenuItem);
                    contextMenuItems.Add(_directReportGetSummaryMenuItem);
                    contextMenuItems.Add(_userGetGroupsMenuItem);
                    contextMenuItems.Add(_userGetSummaryMenuItem);
                    break;
                case QueryType.ContextUserGroups:
                    contextMenuItems.Add(_groupGetComputersMenuItem);
                    contextMenuItems.Add(_groupGetUsersMenuItem);
                    contextMenuItems.Add(_groupGetUsersDirectReportsMenuItem);
                    contextMenuItems.Add(_groupGetUsersGroupsMenuItem);
                    contextMenuItems.Add(_userGetDirectReportsMenuItem);
                    contextMenuItems.Add(_userGetSummaryMenuItem);
                    break;
                case QueryType.OuComputers:
                    contextMenuItems.Add(_computerGetGroupsMenuItem);
                    contextMenuItems.Add(_computerGetSummaryMenuItem);
                    break;
                case QueryType.OuGroups:
                    contextMenuItems.Add(_groupGetComputersMenuItem);
                    contextMenuItems.Add(_groupGetUsersMenuItem);
                    contextMenuItems.Add(_groupGetUsersDirectReportsMenuItem);
                    contextMenuItems.Add(_groupGetUsersGroupsMenuItem);
                    contextMenuItems.Add(_groupGetSummaryMenuItem);
                    break;
                case QueryType.OuUsers:
                    contextMenuItems.Add(_userGetDirectReportsMenuItem);
                    contextMenuItems.Add(_userGetGroupsMenuItem);
                    contextMenuItems.Add(_userGetSummaryMenuItem);
                    break;
                case QueryType.OuUsersDirectReports:
                    contextMenuItems.Add(
                        _directReportGetDirectReportsMenuItem);
                    contextMenuItems.Add(_directReportGetGroupsMenuItem);
                    contextMenuItems.Add(_directReportGetSummaryMenuItem);
                    contextMenuItems.Add(_userGetDirectReportsMenuItem);
                    contextMenuItems.Add(_userGetGroupsMenuItem);
                    contextMenuItems.Add(_userGetSummaryMenuItem);
                    break;
                case QueryType.OuUsersGroups:
                    contextMenuItems.Add(_groupGetComputersMenuItem);
                    contextMenuItems.Add(_groupGetUsersMenuItem);
                    contextMenuItems.Add(_groupGetUsersDirectReportsMenuItem);
                    contextMenuItems.Add(_groupGetUsersGroupsMenuItem);
                    contextMenuItems.Add(_userGetDirectReportsMenuItem);
                    contextMenuItems.Add(_userGetGroupsMenuItem);
                    contextMenuItems.Add(_userGetSummaryMenuItem);
                    break;
                case QueryType.ContextComputerSummary:
                    contextMenuItems.Add(_computerGetGroupsMenuItem);
                    break;
                case QueryType.ContextDirectReportSummary:
                    contextMenuItems.Add(
                        _directReportGetDirectReportsMenuItem);
                    contextMenuItems.Add(_directReportGetGroupsMenuItem);
                    break;
                case QueryType.ContextGroupSummary:
                    contextMenuItems.Add(_groupGetComputersMenuItem);
                    contextMenuItems.Add(_groupGetUsersMenuItem);
                    contextMenuItems.Add(_groupGetUsersDirectReportsMenuItem);
                    contextMenuItems.Add(_groupGetUsersGroupsMenuItem);
                    break;
                case QueryType.ContextUserSummary:
                    contextMenuItems.Add(_userGetDirectReportsMenuItem);
                    contextMenuItems.Add(_userGetGroupsMenuItem);
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
            return contextMenuItems;
        }

        private async void GetContextComputerGroupsCommandExecute()
        {
            await
                RunQuery(QueryType.ContextComputerGroups,
                    GetSelectedComputerDistinguishedName());
        }

        private async void GetContextComputerSummaryCommandExecute()
        {
            await RunQuery(
                QueryType.ContextComputerSummary,
                GetSelectedComputerDistinguishedName());
        }

        private async void GetContextDirectReportDirectReportsCommandExecute()
        {
            await
                RunQuery(QueryType.ContextDirectReportDirectReports,
                    GetSelectedGroupDistinguishedName());
        }

        private async void GetContextDirectReportGroupsCommandExecute()
        {
            await
                RunQuery(QueryType.ContextDirectReportGroups,
                    GetSelectedDirectReportDistinguishedName());
        }

        private async void GetContextGroupComputersCommandExecute()
        {
            await
                RunQuery(QueryType.ContextGroupComputers,
                    GetSelectedGroupDistinguishedName());
        }

        private async void GetContextGroupUsersCommandExecute()
        {
            await RunQuery(QueryType.ContextGroupUsers);
        }

        private async void GetContextGroupUsersDirectReportsCommandExecute()
        {
            await
                RunQuery(QueryType.ContextGroupUsersDirectReports,
                    GetSelectedGroupDistinguishedName());
        }

        private async void GetContextGroupUsersGroupsCommandExecute()
        {
            await
                RunQuery(QueryType.ContextGroupUsersGroups,
                    GetSelectedGroupDistinguishedName());
        }

        private async void GetContextUserDirectReportsCommandExecute()
        {
            await
                RunQuery(QueryType.ContextUserDirectReports,
                    GetSelectedUserDistinguishedName());
        }

        private async void GetContextUserGroupsCommandExecute()
        {
            await
                RunQuery(QueryType.ContextUserGroups,
                    GetSelectedUserDistinguishedName());
        }

        private async void GetOuComputersCommandExecute()
        {
            await RunQuery(QueryType.OuComputers);
        }

        private async void GetOuGroupsCommandExecute()
        {
            await RunQuery(QueryType.OuGroups);
        }

        private async void GetOuUsersCommandExecute()
        {
            await RunQuery(QueryType.OuUsers);
        }

        private async void GetOuUsersDirectReportsCommandExecute()
        {
            await RunQuery(QueryType.OuUsersDirectReports);
        }

        private async void GetOuUsersGroupsCommandExecute()
        {
            await RunQuery(QueryType.OuUsersGroups);
        }

        private async void GetContextUserSummaryCommandExecute()
        {
            await RunQuery(
                QueryType.ContextUserSummary,
                GetSelectedUserDistinguishedName());
        }

        private async void GetContextDirectReportSummaryCommandExecute()
        {
            await RunQuery(
                QueryType.ContextDirectReportSummary,
                GetSelectedDirectReportDistinguishedName());
        }

        private async void GetContextGroupSummaryCommandExecute()
        {
            await RunQuery(
                QueryType.ContextGroupSummary,
                GetSelectedGroupDistinguishedName());
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

        private void HideMessage()
        {
            MessageContent = string.Empty;
            MessageVisibility = Visibility.Hidden;
        }

        private void NotifyPropertyChanged(
            [CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this,
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

        private async Task RunQuery(QueryType queryType,
            string selectedItemDistinguishedName = null)
        {
            StartTask();
            try
            {
                if(Query != null)
                    PreviousQueries.Push(Query);
                Query = new ActiveDirectoryQuery(
                    queryType, CurrentScope, selectedItemDistinguishedName);
                await Query.Execute();
                Data = Query.Data.ToDataTable().AsDataView();
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

        private void ResetQuery()
        {
            if (PreviousQueries.Any() && PreviousQueries.Peek() != null)
                Query = PreviousQueries.Pop();
        }

        private void CancelCommandExecute()
        {
            _activeDirectoryQuery?.Cancel();
        }

        private bool CancelCommandCanExecute()
        {
            return _activeDirectoryQuery != null;
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

            GetContextUserGroupsCommand = new RelayCommand(
                GetContextUserGroupsCommandExecute);

            GetContextComputerGroupsCommand = new RelayCommand(
                GetContextComputerGroupsCommandExecute);

            GetContextDirectReportDirectReportsCommand = new RelayCommand(
                GetContextDirectReportDirectReportsCommandExecute);

            GetContextDirectReportGroupsCommand = new RelayCommand(
                GetContextDirectReportGroupsCommandExecute);

            GetContextGroupComputersCommand = new RelayCommand(
                GetContextGroupComputersCommandExecute);

            GetContextGroupUsersCommand = new RelayCommand(
                GetContextGroupUsersCommandExecute);

            GetContextGroupUsersDirectReportsCommand = new RelayCommand(
                GetContextGroupUsersDirectReportsCommandExecute);

            GetContextUserDirectReportsCommand = new RelayCommand(
                GetContextUserDirectReportsCommandExecute);

            GetContextGroupUsersGroupsCommand = new RelayCommand(
                GetContextGroupUsersGroupsCommandExecute);

            GetContextDirectReportSummaryCommand = new RelayCommand(
                GetContextDirectReportSummaryCommandExecute);

            GetContextUserSummaryCommand = new RelayCommand(
                GetContextUserSummaryCommandExecute);

            GetContextComputerSummaryCommand = new RelayCommand(
                GetContextComputerSummaryCommandExecute);

            GetContextGroupSummaryCommand = new RelayCommand(
                GetContextGroupSummaryCommandExecute);

            CancelCommand= new RelayCommand(
                CancelCommandExecute, CancelCommandCanExecute);
        }

        private void SetViewVariables()
        {
            ProgressBarVisibility = Visibility.Hidden;
            MessageVisibility = Visibility.Hidden;
            CancelButtonVisibility = Visibility.Hidden;
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
            await Task.Run(() =>
            {
                var fileWriter = new DataFileWriter
                {
                    Data = Data,
                    Scope = _activeDirectoryQuery.Scope,
                    QueryType = _activeDirectoryQuery.QueryType
                };
                ShowMessage("Wrote data to:\n" + fileWriter.WriteToCsv());
            });
            FinishTask();
        }
    }
}