using System;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace MichaelBrandonMorris.ActiveDirectoryTool
{
    /// <summary>
    ///     Class ScopeFetcher.
    /// </summary>
    /// TODO Edit XML Comment Template for ScopeFetcher
    internal class ScopeFetcher
    {
        /// <summary>
        ///     The directory searcher ou filter
        /// </summary>
        /// TODO Edit XML Comment Template for DirectorySearcherOuFilter
        private const string DirectorySearcherOuFilter =
            "(objectCategory=organizationalUnit)";

        /// <summary>
        ///     The LDAP prefix
        /// </summary>
        /// TODO Edit XML Comment Template for LdapPrefix
        private const string LdapPrefix = "LDAP://";

        /// <summary>
        ///     The page size
        /// </summary>
        /// TODO Edit XML Comment Template for PageSize
        private const int PageSize = 1;

        /// <summary>
        ///     Initializes a new instance of the
        ///     <see cref="ScopeFetcher" /> class.
        /// </summary>
        /// TODO Edit XML Comment Template for #ctor
        internal ScopeFetcher()
        {
            using (var principalContext =
                new PrincipalContext(ContextType.Domain))
            {
                using (var directoryEntry =
                    new DirectoryEntry(principalContext.ConnectedServer))
                {
                    Scope = new Scope
                    {
                        Name = directoryEntry.Path,
                        Path = LdapPrefix + directoryEntry.Path
                    };
                }
            }

            FetchScopeList();
        }

        /// <summary>
        ///     Gets or sets the scope.
        /// </summary>
        /// <value>The scope.</value>
        /// TODO Edit XML Comment Template for Scope
        internal Scope Scope
        {
            get;
            set;
        }

        /// <summary>
        ///     Fetches the scope list.
        /// </summary>
        /// TODO Edit XML Comment Template for FetchScopeList
        private void FetchScopeList()
        {
            using (var directorySearcher = new DirectorySearcher(Scope.Path))
            {
                directorySearcher.Filter = DirectorySearcherOuFilter;
                directorySearcher.PageSize = PageSize;

                foreach (SearchResult result in directorySearcher.FindAll())
                {
                    Scope.AddDirectoryScope(
                        new OrganizationalUnit
                        {
                            Path = result.Path
                        });
                }
            }

            Scope.Children.Sort(
                (a, b) => string.Compare(
                    a.Name,
                    b.Name,
                    StringComparison.Ordinal));
        }
    }
}