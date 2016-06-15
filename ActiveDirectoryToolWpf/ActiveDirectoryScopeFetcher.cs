using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.Design.Serialization;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;

namespace ActiveDirectoryToolWpf
{
    public class ActiveDirectoryScopeFetcher
    {
        public ActiveDirectoryScope Scope { get; set; }

        internal ActiveDirectoryScopeFetcher()
        {
            var rootPrincipalContext = new PrincipalContext(
                ContextType.Domain);
            var rootDirectoryEntry = new DirectoryEntry(
                rootPrincipalContext.ConnectedServer);
            Scope = new ActiveDirectoryScope
            {
                Name = rootDirectoryEntry.Path,
                Path = "LDAP://" + rootDirectoryEntry.Path
            };
            FetchScopeList();
        }

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
            Scope.Children.Sort((a,b) => a.Name.CompareTo(b.Name));
        }
    }
}