using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using GalaSoft.MvvmLight.CommandWpf;

namespace ActiveDirectoryToolWpf
{
    public enum QueryType
    {
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
            DefaultGroupAttributes =
            {
                ActiveDirectoryAttribute.GroupSamAccountName,
                ActiveDirectoryAttribute.GroupManagedBy,
                ActiveDirectoryAttribute.GroupDescription,
                ActiveDirectoryAttribute.GroupDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultGroupUsersAttributes =
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
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.GroupSamAccountName,
                ActiveDirectoryAttribute.GroupManagedBy,
                ActiveDirectoryAttribute.GroupDescription,
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
            DefaultUserGroupsAttributes =
            {
                ActiveDirectoryAttribute.UserSamAccountName,
                ActiveDirectoryAttribute.GroupSamAccountName,
                ActiveDirectoryAttribute.UserName,
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.GroupDistinguishedName
            };

        private ActiveDirectorySearcher _activeDirectorySearcher;
        private DataView _data;
        private DataPreparer _dataPreparer;

        private string _messageContent;

        private Visibility _messageVisibility;
        private Visibility _progressBarVisibility;

        private QueryType _queryType;
        private bool _viewIsEnabled;

        public ActiveDirectoryToolViewModel()
        {
            RootScope = new ActiveDirectoryScopeFetcher().Scope;
            SetViewVariables();
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

        public ICommand GetOuUsersGroupsCommand
        {
            get;
            private set;
        }

        public ICommand GetOuUsersDirectReportsCommand
        {
            get;
            private set;
        }

        public ICommand WriteToFileCommand
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

        public QueryType QueryType
        {
            get { return _queryType; }
            set
            {
                _queryType = value;
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

        public bool ViewIsEnabled
        {
            get { return _viewIsEnabled; }
            private set
            {
                _viewIsEnabled = value;
                NotifyPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

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

        private void ShowMessage(string messageContent)
        {
            MessageContent = messageContent + "\n\nDouble-click to dismiss.";
            MessageVisibility = Visibility.Visible;
        }

        private void HideMessage()
        {
            MessageContent = string.Empty;
            MessageVisibility = Visibility.Hidden;
        }

        private void FinishTask()
        {
            ProgressBarVisibility = Visibility.Hidden;
            ViewIsEnabled = true;
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

        private bool WriteToFileCommandCanExecute()
        {
            return (Data != null && Data.Count > 0);
        }

        private void NotifyPropertyChanged(
            [CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(
                this, new PropertyChangedEventArgs(propertyName));
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
        }

        private void SetViewVariables()
        {
            ProgressBarVisibility = Visibility.Hidden;
            MessageVisibility = Visibility.Hidden;
            ViewIsEnabled = true;
            SetUpCommands();
        }
    }
}