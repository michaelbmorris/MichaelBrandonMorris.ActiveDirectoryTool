using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace ActiveDirectoryToolWpf
{
    public static class PrincipalExtensions
    {
        public static DirectoryEntry GetAsDirectoryEntry(this Principal principal)
        {
            return principal.GetUnderlyingObject() as DirectoryEntry;
        }

        public static object GetProperty(this Principal principal, string propertyName)
        {
            return principal.GetAsDirectoryEntry().Properties[propertyName].Value;
        }

        public static string GetPropertyAsString(this Principal principal, string propertyName)
        {
            return principal.GetProperty(propertyName) == null
                ? string.Empty
                : principal.GetProperty(propertyName).ToString();
        }
    }
}