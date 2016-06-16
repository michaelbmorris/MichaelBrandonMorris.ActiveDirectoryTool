using System;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace ActiveDirectoryToolWpf
{
    public class ActiveDirectoryToolViewModel
    {
        private readonly ActiveDirectoryAttribute[] _defaultUserAttributes =
        {
            ActiveDirectoryAttribute.Surname,
            ActiveDirectoryAttribute.GivenName,
            ActiveDirectoryAttribute.UserDisplayName,
            ActiveDirectoryAttribute.UserSamAccountName,
            ActiveDirectoryAttribute.IsActive,
            ActiveDirectoryAttribute.IsAccountLockedOut,
            ActiveDirectoryAttribute.UserDescription,
            ActiveDirectoryAttribute.Title,
            ActiveDirectoryAttribute.Company,
            ActiveDirectoryAttribute.Manager,
            ActiveDirectoryAttribute.UserHomeDrive,
            ActiveDirectoryAttribute.UserHomeDirectory,
            ActiveDirectoryAttribute.UserScriptPath,
            ActiveDirectoryAttribute.EmailAddress,
            ActiveDirectoryAttribute.StreetAddress,
            ActiveDirectoryAttribute.City,
            ActiveDirectoryAttribute.State,
            ActiveDirectoryAttribute.VoiceTelephoneNumber,
            ActiveDirectoryAttribute.Pager,
            ActiveDirectoryAttribute.Mobile,
            ActiveDirectoryAttribute.Fax,
            ActiveDirectoryAttribute.Voip,
            ActiveDirectoryAttribute.Sip,
            ActiveDirectoryAttribute.UserUserPrincipalName,
            ActiveDirectoryAttribute.UserDistinguishedName
        };

        private readonly ActiveDirectoryAttribute[] _defaultUserGroupsAttributes =
        {
            ActiveDirectoryAttribute.UserSamAccountName,
            ActiveDirectoryAttribute.GroupSamAccountName, 
            ActiveDirectoryAttribute.UserName, 
            ActiveDirectoryAttribute.UserDistinguishedName
        };

        private readonly IActiveDirectoryToolView _view;
        private ActiveDirectorySearcher _searcher;
        private DataPreparer _dataPreparer;

        public ActiveDirectoryToolViewModel(IActiveDirectoryToolView view)
        {
            _view = view;
            _view.GetUsersClicked += OnGetUsers;
            _view.GetUsersGroupsClicked += OnGetUsersGroups;
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