using System;
using System.Data;

namespace ActiveDirectoryToolWpf
{
    public interface IActiveDirectoryToolView
    {
        ActiveDirectoryScope Scope { get; }
        event Action GetComputersClicked;
        event Action GetDirectReportsClicked;
        event Action GetGroupsClicked;
        event Action GetUsersClicked;
        event Action GetUsersGroupsClicked;
        void SetDataGridData(DataView dataView);
        void ShowMessage(string message);
        void ToggleEnabled();
    }
}