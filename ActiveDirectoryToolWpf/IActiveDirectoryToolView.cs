using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace ActiveDirectoryToolWpf
{
    public interface IActiveDirectoryToolView
    {
        event Action GetUsersClicked;
        event Action GetUsersGroupsClicked;
        event Action GetDirectReportsClicked;
        event Action GetComputersClicked;
        event Action GetGroupsClicked;
        ActiveDirectoryScope Scope { get; }
        void SetDataGridData(DataView dataView);
        void ShowMessage(string message);
        void ToggleEnabled();
    }
}
