using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Linq;

namespace ActiveDirectoryToolWpf
{
    public class DirectReports
    {
        public IEnumerable<UserPrincipal> Reports { get; set; }
        public UserPrincipal User { get; set; }
    }

    public class UserGroups
    {
        public IEnumerable<GroupPrincipal> Groups { get; set; }
        public UserPrincipal User { get; set; }
    }

    public class ActiveDirectorySearcher
    {
        public ActiveDirectorySearcher(ActiveDirectoryScope scope)
        {
            Scope = scope;
        }

        private ActiveDirectoryScope Scope { get; }

        private PrincipalContext PrincipalContext => new PrincipalContext(
            ContextType.Domain, Scope.Domain, Scope.Context);

        public static IEnumerable<ComputerPrincipal> GetComputersFromContext(
            PrincipalContext context)
        {
            IEnumerable<ComputerPrincipal> computers;
            using (var searcher = new PrincipalSearcher(
                new ComputerPrincipal(context)))
            {
                computers =
                    searcher.FindAll().OfType<ComputerPrincipal>().ToList();
            }
            return computers;
        }

        public static IEnumerable<ComputerPrincipal> GetComputersFromGroup(
            GroupPrincipal group)
        {
            return group.GetMembers().OfType<ComputerPrincipal>();
        }

        public static IEnumerable<DirectReports> GetDirectReportsFromUsers(
            IEnumerable<UserPrincipal> users)
        {
            return users.Select(user => new DirectReports
            {
                User = user,
                Reports = user.GetDirectReports()
            }).ToList();
        }

        public static IEnumerable<DirectReports> GetDirectReportsFromContext(
            PrincipalContext context)
        {
            return GetDirectReportsFromUsers(GetUsersFromContext(context));
        }

        public static IEnumerable<GroupPrincipal> GetGroupsFromContext(
            PrincipalContext context)
        {
            return GetGroupsFromUsers(GetUsersFromContext(context));
        }

        public static IEnumerable<GroupPrincipal> GetGroupsFromUsers(
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

        public static IEnumerable<UserGroups> GetUserGroupsFromContext(
            PrincipalContext context)
        {
            return GetUserGroupsFromUsers(GetUsersFromContext(context));
        }

        public static IEnumerable<UserGroups> GetUserGroupsFromUsers(
            IEnumerable<UserPrincipal> users)
        {
            return users.Select(user => new UserGroups
            {
                User = user,
                Groups = user.GetGroups().OfType<GroupPrincipal>().ToList()
            }).ToList();
        }

        public static IEnumerable<UserPrincipal> GetUsersFromContext(
            PrincipalContext context)
        {
            IEnumerable<UserPrincipal> users;
            using (var searcher = new PrincipalSearcher(new UserPrincipal(
                context)))
            {
                users = searcher.FindAll().OfType<UserPrincipal>().ToList();
            }
            return users;
        }

        public IEnumerable<UserPrincipal> GetUsers()
        {
            return GetUsersFromContext(PrincipalContext);
        }

        public IEnumerable<UserGroups> GetUsersGroups()
        {
            return GetUserGroupsFromContext(PrincipalContext);
        }

        public IEnumerable<DirectReports> GetDirectReports()
        {
            return GetDirectReportsFromContext(PrincipalContext);
        }

        public static IEnumerable<UserPrincipal> GetUsersFromGroup(
            GroupPrincipal group)
        {
            return group.GetMembers().OfType<UserPrincipal>().ToList();
        }
    }
}