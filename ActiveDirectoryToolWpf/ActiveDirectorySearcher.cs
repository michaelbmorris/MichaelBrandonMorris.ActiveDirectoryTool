using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiveDirectoryToolWpf
{
    internal struct DirectReports
    {
        internal IEnumerable<UserPrincipal> Reports { get; set; }
        internal UserPrincipal User { get; set; }
    }

    internal struct UserGroups
    {
        internal IEnumerable<GroupPrincipal> Groups { get; set; }
        internal UserPrincipal User { get; set; }
    }

    internal static class UserPrincipalExtensions
    {
        internal static DirectoryEntry GetAsDirectoryEntry(
            this UserPrincipal user)
        {
            return user.GetUnderlyingObject() as DirectoryEntry;
        }

        internal static IEnumerable<string> GetDirectReportDistinguishedNames(
            this UserPrincipal user)
        {
            return ((IEnumerable<string>)user.GetProperty("directReports"));
        }

        internal static IEnumerable<UserPrincipal> GetDirectReports(
            this UserPrincipal user)
        {
            var directReports = new List<UserPrincipal>();
            foreach (var directReportDistinguishedName in
                user.GetDirectReportDistinguishedNames())
            {
                directReports.Add(UserPrincipal.FindByIdentity(
                    user.Context,
                    IdentityType.DistinguishedName,
                    directReportDistinguishedName));
            }
            return directReports;
        }

        internal static object GetProperty(
            this UserPrincipal user, string propertyName)
        {
            return user.GetAsDirectoryEntry().Properties[propertyName];
        }
    }

    internal class ActiveDirectorySearcher
    {
        internal ActiveDirectoryScope Scope { get; set; }

        internal static IEnumerable<ComputerPrincipal> GetComputersFromContext(
            PrincipalContext context)
        {
            var computers = new List<ComputerPrincipal>();
            using (var searcher = new PrincipalSearcher(
                    new ComputerPrincipal(context)))
            {
                computers =
                    searcher.FindAll().OfType<ComputerPrincipal>().ToList();
            }
            return computers;
        }

        internal static IEnumerable<ComputerPrincipal> GetComputersFromGroup(
            GroupPrincipal group)
        {
            return group.GetMembers().OfType<ComputerPrincipal>();
        }

        internal static IEnumerable<DirectReports> GetDirectReportsFromUsers(
                    IEnumerable<UserPrincipal> users)
        {
            var directReports = new List<DirectReports>();
            foreach (var user in users)
            {
                directReports.Add(new DirectReports
                {
                    User = user,
                    Reports = user.GetDirectReports()
                });
            }
            return directReports;
        }

        internal static IEnumerable<DirectReports> GetDiretReportsFromContext(
            PrincipalContext context)
        {
            return GetDirectReportsFromUsers(GetUsersFromContext(context));
        }

        internal static IEnumerable<GroupPrincipal> GetGroupsFromContext(
            PrincipalContext context)
        {
            return GetGroupsFromUsers(GetUsersFromContext(context));
        }

        internal static IEnumerable<GroupPrincipal> GetGroupsFromUsers(
            IEnumerable<UserPrincipal> users)
        {
            var groups = new HashSet<GroupPrincipal>();
            foreach (var user in users)
            {
                groups.UnionWith(
                    user.GetGroups().OfType<GroupPrincipal>().ToList());
            }

            return groups;
        }

        internal static IEnumerable<UserGroups> GetUserGroupsFromContext(
            PrincipalContext context)
        {
            return GetUserGroupsFromUsers(GetUsersFromContext(context));
        }

        internal static IEnumerable<UserGroups> GetUserGroupsFromUsers(
            IEnumerable<UserPrincipal> users)
        {
            var userGroups = new List<UserGroups>();
            foreach (var user in users)
            {
                userGroups.Add(new UserGroups
                {
                    User = user,
                    Groups = user.GetGroups().OfType<GroupPrincipal>().ToList()
                });
            }
            return userGroups;
        }

        internal static IEnumerable<UserPrincipal> GetUsersFromContext(
            PrincipalContext context)
        {
            var users = new List<UserPrincipal>();
            using (var searcher = new PrincipalSearcher(new UserPrincipal(
                context)))
            {
                users = searcher.FindAll().OfType<UserPrincipal>().ToList();
            }
            return users;
        }

        internal IEnumerable<UserPrincipal> GetUsers()
        {
            Debug.WriteLine(Scope.Path);
            var principalContext = new PrincipalContext(ContextType.Domain, Scope.Domain, Scope.Path.Replace("LDAP://", string.Empty));
            var users = new List<UserPrincipal>();
            using (var searcher = new PrincipalSearcher(new UserPrincipal(principalContext)))
            {
                users = searcher.FindAll().OfType<UserPrincipal>().ToList();
            }
            return users;
        }

        internal static IEnumerable<UserPrincipal> GetUsersFromGroup(
            GroupPrincipal group)
        {
            return group.GetMembers().OfType<UserPrincipal>().ToList();
        }
    }
}
