using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiveDirectoryToolWpf
{
    public static class GroupPrincipalExtensions
    {
        public const string ManagedBy = "managedBy";

        internal static string GetManagedBy(this GroupPrincipal group)
        {
            return group.GetPropertyAsString(ManagedBy);
        }
    }
}
