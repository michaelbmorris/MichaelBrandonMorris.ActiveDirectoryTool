using System.Data;
using System.Linq;

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

        private void OnGetDirectReports()
        {
            _searcher = new ActiveDirectorySearcher(_view.Scope);
            _dataPreparer = new DataPreparer
            {
                Data = _searcher.GetDirectReports(),
                Attributes = _defaultDirectReportsAttributes.ToList()
            };
            _view.SetDataGridData(
                _dataPreparer.GetResults().ToDataTable().AsDataView());
        }

        private void OnGetGroups()
        {
            _searcher = new ActiveDirectorySearcher(_view.Scope);
            _dataPreparer = new DataPreparer
            {
                Data = _searcher.GetGroups(),
                Attributes = _defaultGroupAttributes.ToList()
            };
            _view.SetDataGridData(
                _dataPreparer.GetResults().ToDataTable().AsDataView());
        }

        private void OnGetUsers()
        {
            _searcher = new ActiveDirectorySearcher(_view.Scope);
            _dataPreparer = new DataPreparer
            {
                Data = _searcher.GetUsers(),
                Attributes = _defaultUserAttributes.ToList()
            };
            _view.SetDataGridData(
                _dataPreparer.GetResults().ToDataTable().AsDataView());
        }

        private void OnGetUsersGroups()
        {
            _searcher = new ActiveDirectorySearcher(_view.Scope);
            _dataPreparer = new DataPreparer
            {
                Data = _searcher.GetUsersGroups(),
                Attributes = _defaultUserGroupsAttributes.ToList()
            };
            _view.SetDataGridData(
                _dataPreparer.GetResults().ToDataTable().AsDataView());
        }
    }
}