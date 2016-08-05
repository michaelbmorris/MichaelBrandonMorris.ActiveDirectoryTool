using System;

namespace MichaelBrandonMorris.ActiveDirectoryTool
{
    public class OrganizationalUnit
    {
        private const char Comma = ',';
        private const string DomainComponentPrefix = ",DC=";
        private const string LdapProtocol = "LDAP://";
        private const string OrganizationalUnitPrefix = "OU=";
        private const int StringStartIndex = 0;

        public string Path { get; set; }

        public string[] Split()
        {
            var namesOnly = Path.Replace(LdapProtocol, string.Empty)
                .Replace(OrganizationalUnitPrefix, string.Empty);
            var firstDomainComponentIndex = namesOnly.IndexOf(
                DomainComponentPrefix, StringComparison.Ordinal);
            namesOnly = namesOnly.Substring(
                StringStartIndex, firstDomainComponentIndex);
            var split = namesOnly.Split(Comma);
            Array.Reverse(split);
            return split;
        }
    }
}