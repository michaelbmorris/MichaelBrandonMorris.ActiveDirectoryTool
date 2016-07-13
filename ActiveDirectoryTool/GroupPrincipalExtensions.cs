using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Linq;

namespace ActiveDirectoryToolWpf
{
    public static class GroupPrincipalExtensions
    {
        public const string ManagedBy = "managedBy";

        internal static string GetManagedBy(this GroupPrincipal groupPrincipal)
        {
            return groupPrincipal.GetPropertyAsString(ManagedBy);
        }

        public static IEnumerable<UserPrincipal> GetUserPrincipals(
            this GroupPrincipal groupPrincipal)
        {
            return groupPrincipal.Members.OfType<UserPrincipal>();
        }

        public static IEnumerable<ComputerPrincipal> GetComputerPrincipals(
            this GroupPrincipal groupPrincipal)
        {
            return groupPrincipal.Members.OfType<ComputerPrincipal>();
        }
    }
}