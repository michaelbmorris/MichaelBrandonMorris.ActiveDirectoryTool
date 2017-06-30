using System;

namespace MichaelBrandonMorris.ActiveDirectoryTool
{
    /// <summary>
    ///     Class OrganizationalUnit.
    /// </summary>
    /// TODO Edit XML Comment Template for OrganizationalUnit
    public class OrganizationalUnit
    {
        /// <summary>
        ///     The comma
        /// </summary>
        /// TODO Edit XML Comment Template for Comma
        private const char Comma = ',';

        /// <summary>
        ///     The domain component prefix
        /// </summary>
        /// TODO Edit XML Comment Template for DomainComponentPrefix
        private const string DomainComponentPrefix = ",DC=";

        /// <summary>
        ///     The LDAP protocol
        /// </summary>
        /// TODO Edit XML Comment Template for LdapProtocol
        private const string LdapProtocol = "LDAP://";

        /// <summary>
        ///     The organizational unit prefix
        /// </summary>
        /// TODO Edit XML Comment Template for OrganizationalUnitPrefix
        private const string OrganizationalUnitPrefix = "OU=";

        /// <summary>
        ///     The string start index
        /// </summary>
        /// TODO Edit XML Comment Template for StringStartIndex
        private const int StringStartIndex = 0;

        /// <summary>
        ///     Gets or sets the path.
        /// </summary>
        /// <value>The path.</value>
        /// TODO Edit XML Comment Template for Path
        public string Path
        {
            get;
            set;
        }

        /// <summary>
        ///     Splits this instance.
        /// </summary>
        /// <returns>System.String[].</returns>
        /// TODO Edit XML Comment Template for Split
        public string[] Split()
        {
            var namesOnly = Path.Replace(LdapProtocol, string.Empty)
                .Replace(OrganizationalUnitPrefix, string.Empty);

            var firstDomainComponentIndex = namesOnly.IndexOf(
                DomainComponentPrefix,
                StringComparison.Ordinal);

            namesOnly = namesOnly.Substring(
                StringStartIndex,
                firstDomainComponentIndex);

            var split = namesOnly.Split(Comma);
            Array.Reverse(split);
            return split;
        }
    }
}