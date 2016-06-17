using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.Linq;

namespace ActiveDirectoryToolWpf
{
    public class ActiveDirectorySearcher
    {
        public ActiveDirectorySearcher(ActiveDirectoryScope scope)
        {
            Scope = scope;
        }

        private PrincipalContext PrincipalContext => new PrincipalContext(
            ContextType.Domain, Scope.Domain, Scope.Context);

        private ActiveDirectoryScope Scope { get; }

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

        public static IEnumerable<DirectReports> GetDirectReportsFromContext(
            PrincipalContext context)
        {
            return GetDirectReportsFromUsers(GetUsersFromContext(context));
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
            var userGroups = new List<UserGroups>();
            foreach (var user in users)
            {
                var groups = new List<GroupPrincipal>();
                try
                {
                    groups.AddRange(user.GetGroups().OfType<GroupPrincipal>());
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
                userGroups.Add(new UserGroups
                {
                    User = user,
                    Groups = groups
                });
            }
            return userGroups;
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

        public static IEnumerable<UserPrincipal> GetUsersFromGroup(
            GroupPrincipal group)
        {
            return group.GetMembers().OfType<UserPrincipal>().ToList();
        }

        public IEnumerable<DirectReports> GetDirectReports()
        {
            return GetDirectReportsFromContext(PrincipalContext);
        }

        public IEnumerable<GroupPrincipal> GetGroups()
        {
            return GetGroupsFromContext(PrincipalContext);
        }

        public IEnumerable<UserPrincipal> GetUsers()
        {
            return GetUsersFromContext(PrincipalContext);
        }

        public IEnumerable<UserGroups> GetUsersGroups()
        {
            return GetUserGroupsFromContext(PrincipalContext);
        }
    }

    public class DirectReports : ExtendedPrincipalBase, IDisposable
    {
        public IEnumerable<UserPrincipal> Reports { get; set; }

        public void Dispose()
        {
            User?.Dispose();
            foreach (var report in Reports)
                report?.Dispose();
        }
    }

    public abstract class ExtendedPrincipalBase
    {
        public UserPrincipal User { get; set; }
    }

    public class UserGroups : ExtendedPrincipalBase, IDisposable
    {
        public IEnumerable<GroupPrincipal> Groups { get; set; }

        public void Dispose()
        {
            User?.Dispose();
            foreach (var group in Groups)
                group?.Dispose();
        }
    }
}