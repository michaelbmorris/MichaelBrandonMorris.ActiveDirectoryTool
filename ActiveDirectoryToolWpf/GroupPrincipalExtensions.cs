using System.DirectoryServices.AccountManagement;

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