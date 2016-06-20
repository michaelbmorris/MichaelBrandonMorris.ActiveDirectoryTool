using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Threading.Tasks;

namespace ActiveDirectoryToolWpf
{
    public class ActiveDirectoryToolViewModel
    {
        private readonly ActiveDirectoryAttribute[]
            _defaultDirectReportsAttributes =
            {
                ActiveDirectoryAttribute.UserDisplayName,
                ActiveDirectoryAttribute.UserSamAccountName,
                ActiveDirectoryAttribute.DirectReportDisplayName,
                ActiveDirectoryAttribute.DirectReportSamAccountName
            };

        private readonly ActiveDirectoryAttribute[] _defaultGroupAttributes =
        {
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
                ActiveDirectoryAttribute.UserDistinguishedName
            };

        private readonly IActiveDirectoryToolView _view;
        private DataPreparer _dataPreparer;
        private ActiveDirectorySearcher _searcher;

        public ActiveDirectoryToolViewModel(IActiveDirectoryToolView view)
        {
            _view = view;
            _view.GetUsersClicked += OnGetUsers;
            _view.GetUsersGroupsClicked += OnGetUsersGroups;
            _view.GetDirectReportsClicked += OnGetDirectReports;
            _view.GetGroupsClicked += OnGetGroups;
        }

        private async void OnGetDirectReports()
        {
            _view.SetDataGridData(null);
            _view.ToggleProgressBarVisibility();
            _searcher = new ActiveDirectorySearcher(_view.Scope);
            List<ExpandoObject> data = null;
            _view.ToggleEnabled();
            await Task.Run(() =>
            {
                _dataPreparer = new DataPreparer
                {
                    Data = _searcher.GetDirectReports(),
                    Attributes = _defaultDirectReportsAttributes.ToList()
                };
                data = _dataPreparer.GetResults();
            });

            _view.ToggleProgressBarVisibility();
            _view.SetDataGridData(data.ToDataTable().AsDataView());
            _view.ToggleEnabled();
        }

        private async void OnGetGroups()
        {
            _view.SetDataGridData(null);
            _view.ToggleProgressBarVisibility();
            _searcher = new ActiveDirectorySearcher(_view.Scope);
            List<ExpandoObject> data = null;
            _view.ToggleEnabled();
            await Task.Run(() =>
            {
                _dataPreparer = new DataPreparer
                {
                    Data = _searcher.GetGroups(),
                    Attributes = _defaultGroupAttributes.ToList()
                };
                data = _dataPreparer.GetResults();
            });

            _view.ToggleProgressBarVisibility();
            _view.SetDataGridData(data.ToDataTable().AsDataView());
            _view.ToggleEnabled();
        }

        private async void OnGetUsers()
        {
            _view.SetDataGridData(null);
            _view.ToggleProgressBarVisibility();
            _searcher = new ActiveDirectorySearcher(_view.Scope);
            List<ExpandoObject> data = null;
            _view.ToggleEnabled();
            await Task.Run(() =>
            {
                _dataPreparer = new DataPreparer
                {
                    Data = _searcher.GetUsers(),
                    Attributes = _defaultUserAttributes.ToList()
                };
                data = _dataPreparer.GetResults();
            });

            _view.ToggleProgressBarVisibility();
            _view.SetDataGridData(data.ToDataTable().AsDataView());
            _view.ToggleEnabled();
        }

        private async void OnGetUsersGroups()
        {
            _view.SetDataGridData(null);
            _view.ToggleProgressBarVisibility();
            _searcher = new ActiveDirectorySearcher(_view.Scope);
            List<ExpandoObject> data = null;
            _view.ToggleEnabled();
            await Task.Run(() =>
            {
                _dataPreparer = new DataPreparer
                {
                    Data = _searcher.GetUsersGroups(),
                    Attributes = _defaultUserGroupsAttributes.ToList()
                };
                data = _dataPreparer.GetResults();
            });

            _view.ToggleProgressBarVisibility();
            _view.SetDataGridData(data.ToDataTable().AsDataView());
            _view.ToggleEnabled();
        }
    }
}