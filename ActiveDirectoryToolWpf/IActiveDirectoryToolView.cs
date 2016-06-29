using System;
using System.Data;

namespace ActiveDirectoryToolWpf
{
    public interface IActiveDirectoryToolView
    {
        ActiveDirectoryScope Scope { get; }
        string SelectedItemDistinguishedName { get; set; }
        string SearchText { get; set; }
        QueryType QueryType { get; set; }
        event Action GetComputersButtonClicked;
        event Action GetDirectReportsButtonClicked;
        event Action GetGroupComputersMenuItemClicked;
        event Action GetGroupsButtonClicked;
        event Action GetGroupUsersMenuItemClicked;
        event Action GetUserDirectReportsMenuItemClicked;
        event Action GetUserGroupsMenuItemClicked;
        event Action GetUsersButtonClicked;
        event Action GetUsersGroupsButtonClicked;
        event Action SearchButtonClicked;
        void GenerateContextMenu();
        void SetDataGridData(DataView dataView);
        void ShowMessage(string message);
        void ToggleEnabled();
        void ToggleProgressBarVisibility();
    }
}