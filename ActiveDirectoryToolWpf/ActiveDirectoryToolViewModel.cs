using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Linq;
using System.Threading.Tasks;

namespace ActiveDirectoryToolWpf
{
    public class ActiveDirectoryToolViewModel
    {
        private const string NoResultsErrorMessage =
            "No results found. Please ensure you are searching for the " +
            "correct principal type in the correct OU.";

        private readonly ActiveDirectoryAttribute[]
            _defaultComputerAttributes =
            {
                ActiveDirectoryAttribute.ComputerName,
                ActiveDirectoryAttribute.ComputerDescription,
                ActiveDirectoryAttribute.ComputerDistinguishedName
            };

        private readonly ActiveDirectoryAttribute[]
            _defaultDirectReportsAttributes =
            {
                ActiveDirectoryAttribute.UserDisplayName,
                ActiveDirectoryAttribute.UserSamAccountName,
                ActiveDirectoryAttribute.DirectReportDisplayName,
                ActiveDirectoryAttribute.DirectReportSamAccountName,
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.DirectReportDistinguishedName
            };

        private readonly ActiveDirectoryAttribute[] _defaultGroupAttributes =
        {
            ActiveDirectoryAttribute.GroupSamAccountName,
            ActiveDirectoryAttribute.GroupManagedBy,
            ActiveDirectoryAttribute.GroupDescription,
            ActiveDirectoryAttribute.GroupDistinguishedName
        };

        private readonly ActiveDirectoryAttribute[]
            _defaultGroupUsersAttributes =
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

        private readonly ActiveDirectoryAttribute[] _defaultUserAttributes =
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

        private readonly ActiveDirectoryAttribute[]
            _defaultUserGroupsAttributes =
            {
                ActiveDirectoryAttribute.UserSamAccountName,
                ActiveDirectoryAttribute.GroupSamAccountName,
                ActiveDirectoryAttribute.UserName,
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.GroupDistinguishedName
            };

        private readonly IActiveDirectoryToolView _view;
        private IEnumerable<ExpandoObject> _data;
        private DataPreparer _dataPreparer;
        private ActiveDirectorySearcher _searcher;

        public ActiveDirectoryToolViewModel(IActiveDirectoryToolView view)
        {
            _view = view;
            _view.GetComputersButtonClicked +=
                OnGetComputersButtonClicked;
            _view.GetDirectReportsButtonClicked +=
                OnGetDirectReportsButtonClicked;

            _view.GetGroupsButtonClicked += OnGetGroupsButtonClicked;
            _view.GetUsersButtonClicked += OnGetUsersButtonClicked;
            _view.GetUsersGroupsButtonClicked += OnGetUsersGroupsButtonClicked;

            _view.GetUserGroupsMenuItemClicked +=
                OnGetUserGroupsMenuItemClicked;

            _view.GetGroupUsersMenuItemClicked +=
                OnGetGroupUsersMenuItemClicked;

            _view.GetUserDirectReportsMenuItemClicked +=
                OnGetUserDirectReportsMenuItemClicked;

            _view.GetGroupComputersMenuItemClicked +=
                OnGetGroupComputersMenuItemClicked;

            _view.SearchButtonClicked += OnSearchButtonClicked;
        }

        private void FinishTask()
        {
            _view.ToggleProgressBarVisibility();
            try
            {
                _view.SetDataGridData(_data.ToDataTable().AsDataView());
            }
            catch (ArgumentNullException)
            {
                _view.ShowMessage(NoResultsErrorMessage);
            }

            _view.GenerateContextMenu();
            _view.ToggleEnabled();
        }

        private async void OnGetGroupComputersMenuItemClicked()
        {
            StartTask();
            await Task.Run(() =>
            {
                var principalContext = new PrincipalContext(
                    ContextType.Domain);
                var groupPrincipal = GroupPrincipal.FindByIdentity(
                    principalContext, _view.SelectedItemDistinguishedName);
                _dataPreparer = new DataPreparer
                {
                    Data = ActiveDirectorySearcher.GetComputersFromGroup(
                        groupPrincipal),
                    Attributes = _defaultComputerAttributes.ToList()
                };
                _data = _dataPreparer.GetResults();
            });

            FinishTask();
        }

        private async void OnGetComputersButtonClicked()
        {
            StartTask();
            await Task.Run(() =>
            {
                _dataPreparer = new DataPreparer
                {
                    Data = _searcher.GetComputers(),
                    Attributes = _defaultComputerAttributes.ToList()
                };
                _data = _dataPreparer.GetResults();
            });

            FinishTask();
        }

        private async void OnGetDirectReportsButtonClicked()
        {
            StartTask();
            await Task.Run(() =>
            {
                _dataPreparer = new DataPreparer
                {
                    Data = _searcher.GetDirectReports(),
                    Attributes = _defaultDirectReportsAttributes.ToList()
                };
                _data = _dataPreparer.GetResults();
            });

            FinishTask();
        }

        private async void OnGetGroupsButtonClicked()
        {
            StartTask();
            await Task.Run(() =>
            {
                _dataPreparer = new DataPreparer
                {
                    Data = _searcher.GetGroups(),
                    Attributes = _defaultGroupAttributes.ToList()
                };
                _data = _dataPreparer.GetResults();
            });

            FinishTask();
        }

        private async void OnGetGroupUsersMenuItemClicked()
        {
            StartTask();
            await Task.Run(() =>
            {
                var principalContext = new PrincipalContext(
                    ContextType.Domain);
                var groupPrincipal = GroupPrincipal.FindByIdentity(
                    principalContext, _view.SelectedItemDistinguishedName);
                _dataPreparer = new DataPreparer
                {
                    Data = ActiveDirectorySearcher.GetUsersFromGroup(
                        groupPrincipal),
                    Attributes = _defaultGroupUsersAttributes.ToList()
                };
                _data = _dataPreparer.GetResults();
            });

            FinishTask();
        }

        private async void OnGetUserDirectReportsMenuItemClicked()
        {
            StartTask();
            await Task.Run(() =>
            {
                var principalContext = new PrincipalContext(
                    ContextType.Domain);
                var userPrincipal = UserPrincipal.FindByIdentity(
                    principalContext, _view.SelectedItemDistinguishedName);
                _dataPreparer = new DataPreparer
                {
                    Data = new[]
                    {
                        ActiveDirectorySearcher.GetDirectReportsFromUser(
                            userPrincipal)
                    },
                    Attributes = _defaultDirectReportsAttributes.ToList()
                };
                _data = _dataPreparer.GetResults();
            });

            FinishTask();
        }

        private async void OnGetUserGroupsMenuItemClicked()
        {
            StartTask();
            await Task.Run(() =>
            {
                var principalContext = new PrincipalContext(
                    ContextType.Domain);
                var userPrincipal = UserPrincipal.FindByIdentity(
                    principalContext, _view.SelectedItemDistinguishedName);
                _dataPreparer = new DataPreparer
                {
                    Data = new[]
                    {
                        ActiveDirectorySearcher.GetUserGroupsFromUser(
                            userPrincipal)
                    },
                    Attributes = _defaultUserGroupsAttributes.ToList()
                };
                _data = _dataPreparer.GetResults();
            });

            FinishTask();
        }

        private async void OnGetUsersButtonClicked()
        {
            StartTask();
            await Task.Run(() =>
            {
                _dataPreparer = new DataPreparer
                {
                    Data = _searcher.GetUsers(),
                    Attributes = _defaultUserAttributes.ToList()
                };
                _data = _dataPreparer.GetResults();
            });

            FinishTask();
        }

        private async void OnGetUsersGroupsButtonClicked()
        {
            StartTask();
            await Task.Run(() =>
            {
                _dataPreparer = new DataPreparer
                {
                    Data = _searcher.GetUsersGroups(),
                    Attributes = _defaultUserGroupsAttributes.ToList()
                };
                _data = _dataPreparer.GetResults();
            });

            FinishTask();
        }

        private void StartTask()
        {
            _view.SetDataGridData(null);
            _view.ToggleProgressBarVisibility();
            _searcher = new ActiveDirectorySearcher(_view.Scope);
            _view.ToggleEnabled();
        }

        private void OnSearchButtonClicked()
        {
            StartTask();
            var directoryEntry = new DirectoryEntry("LDAP://" + _view.Scope.Context);
            var directorySearcher = new DirectorySearcher(directoryEntry);
            var results = directorySearcher.FindAll();
            foreach (SearchResult result in results)
            {
                var de = result.GetDirectoryEntry();
                //Debug.WriteLine(de.SchemaClassName);
                Debug.WriteLine(de.Properties["name"].Value);
            }
            FinishTask();
        }
    }
}