using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace ActiveDirectoryTool
{
    public class ScopeFetcher
    {
        internal ScopeFetcher()
        {
            var rootPrincipalContext = new PrincipalContext(
                ContextType.Domain);
            var rootDirectoryEntry = new DirectoryEntry(
                rootPrincipalContext.ConnectedServer);
            Scope = new Scope
            {
                Name = rootDirectoryEntry.Path,
                Path = "LDAP://" + rootDirectoryEntry.Path
            };
            FetchScopeList();
        }

        public Scope Scope { get; set; }

        private void FetchScopeList()
        {
            using (var directorySearcher = new DirectorySearcher(Scope.Path))
            {
                directorySearcher.Filter =
                    "(objectCategory=organizationalUnit)";
                directorySearcher.PageSize = 1;
                foreach (SearchResult result in directorySearcher.FindAll())
                {
                    Scope.AddDirectoryScope(new OrganizationalUnit
                    {
                        Path = result.Path
                    });
                }
            }
            Scope.Children.Sort((a, b) => a.Name.CompareTo(b.Name));
        }
    }
}