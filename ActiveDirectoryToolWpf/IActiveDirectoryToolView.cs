using System;
using System.Data;

namespace ActiveDirectoryToolWpf
{
    public interface IActiveDirectoryToolView
    {
        ActiveDirectoryScope Scope { get; }
        string SelectedItemDistinguishedName { get; set; }
        event Action GetComputersClicked;
        event Action GetDirectReportsClicked;
        event Action GetGroupComputersClicked;
        event Action GetGroupsClicked;
        event Action GetGroupUsersClicked;
        event Action GetUserDirectReportsClicked;
        event Action GetUserGroupsClicked;
        event Action GetUsersClicked;
        event Action GetUsersGroupsClicked;
        void GenerateContextMenu();
        void SetDataGridData(DataView dataView);
        void ShowMessage(string message);
        void ToggleEnabled();
        void ToggleProgressBarVisibility();
    }
}