using PrimitiveExtensions;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;

namespace ActiveDirectoryToolWpf
{ 

    public static class PrincipalExtensions
    {
        private static readonly char[] LegalCharacters =
        {
            ' ',
            '\\',
            '/',
            ',',
            '.'
        };

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
                : principal.GetProperty(propertyName).ToString().Clean();
        }

        private static string Clean(this string s)
        {
            return new string(s.Where(c => c.IsLetterOrDigit() ||
            c.EqualsAny(LegalCharacters)).ToArray());
        }
    }
}