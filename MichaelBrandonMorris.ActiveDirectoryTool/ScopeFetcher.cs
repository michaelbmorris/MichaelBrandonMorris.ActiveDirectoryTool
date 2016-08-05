using System;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace MichaelBrandonMorris.ActiveDirectoryTool
{
    internal class ScopeFetcher
    {
        private const string LdapPrefix = "LDAP://";

        private const string DirectorySearcherOuFilter =
            "(objectCategory=organizationalUnit)";

        private const int PageSize = 1;

        internal ScopeFetcher()
        {
            using (var principalContext =
                new PrincipalContext(ContextType.Domain))
            {
                using (var directoryEntry = new DirectoryEntry(
                    principalContext.ConnectedServer))
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

        internal Scope Scope { get; set; }

        private void FetchScopeList()
        {
            using (var directorySearcher = new DirectorySearcher(Scope.Path))
            {
                directorySearcher.Filter = DirectorySearcherOuFilter;
                directorySearcher.PageSize = PageSize;
                foreach (SearchResult result in directorySearcher.FindAll())
                {
                    Scope.AddDirectoryScope(new OrganizationalUnit
                    {
                        Path = result.Path
                    });
                }
            }
            Scope.Children.Sort((a, b) => string.Compare(
                    a.Name, b.Name, StringComparison.Ordinal));
        }
    }
}