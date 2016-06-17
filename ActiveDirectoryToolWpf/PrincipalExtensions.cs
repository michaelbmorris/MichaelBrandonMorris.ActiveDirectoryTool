using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace ActiveDirectoryToolWpf
{
    public static class PrincipalExtensions
    {
        public static DirectoryEntry GetAsDirectoryEntry(
            this Principal principal)
        {
            return principal.GetUnderlyingObject() as DirectoryEntry;
        }

        public static PropertyValueCollection GetProperty(
            this Principal principal, string propertyName)
        {
            return principal.GetAsDirectoryEntry().Properties[propertyName];
        }

        public static string GetPropertyAsString(
            this Principal principal, string propertyName)
        {
            return principal.GetProperty(propertyName).Value == null
                ? string.Empty
                : principal.GetProperty(propertyName).Value.ToString();
        }
    }
}