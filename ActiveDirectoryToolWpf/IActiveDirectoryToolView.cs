using System;
using System.Data;

namespace ActiveDirectoryToolWpf
{
    public interface IActiveDirectoryToolView
    {
        ActiveDirectoryScope Scope { get; }
        event Action GetUsersClicked;

        event Action GetUsersGroupsClicked;

        event Action GetDirectReportsClicked;

        event Action GetComputersClicked;

        event Action GetGroupsClicked;

        void SetDataGridData(DataView dataView);

        void ShowMessage(string message);

        void ToggleEnabled();
    }
}