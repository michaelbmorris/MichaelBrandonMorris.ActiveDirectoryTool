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

        private ActiveDirectoryScope Scope
        {
            get;
        }

        public static ComputerGroups GetComputerGroups(
            ComputerPrincipal computer)
        {
            return new ComputerGroups
            {
                Computer = computer,
                Groups = computer.GetGroups().OfType<GroupPrincipal>().ToList()
            };
        }

        public static IEnumerable<UserDirectReports> GetUsersDirectReports(
            GroupPrincipal groupPrincipal)
        {
            var usersDirectReports = new List<UserDirectReports>();
            foreach (var user in groupPrincipal.Members.OfType<UserPrincipal>()
                )
            {
                usersDirectReports.Add(GetUserDirectReports(user));
            }
            return usersDirectReports;
        }

        public static IEnumerable<ComputerGroups> GetComputersGroups(
            PrincipalContext principalContext)
        {
            var computersGroups = new List<ComputerGroups>();
            using (var searcher =
                new PrincipalSearcher(new ComputerPrincipal(principalContext)))
            {
                computersGroups.AddRange(
                    searcher.FindAll().OfType<ComputerPrincipal>().Select(
                        GetComputerGroups));
            }
            return computersGroups;
        }

        public static IEnumerable<ComputerPrincipal> GetComputersFromContext(
            PrincipalContext context)
        {
            IEnumerable<ComputerPrincipal> computers;
            using (var searcher = new PrincipalSearcher(
                new ComputerPrincipal(context)))
            {
                computers = searcher.FindAll().OfType<ComputerPrincipal>()
                    .ToList();
            }
            return computers;
        }

        public static IEnumerable<ComputerPrincipal> GetComputersFromGroup(
            GroupPrincipal group)
        {
            return group.GetMembers().OfType<ComputerPrincipal>();
        }

        /*public static IEnumerable<UserGroups> GetUsersGroups(
            GroupPrincipal groupPrincipal)
        {
            foreach (var user in groupPrincipal.Members.OfType<UserPrincipal>()
                )
            {
                
            }
        }

        public static IEnumerable<GroupUsersGroups> GetGroupUsersGroups(
            GroupPrincipal groupPrincipal)
        {
            var groupUsersGroups = new List<GroupUsersGroups>();
            foreach (var user in groupPrincipal.Members.OfType<UserPrincipal>()
                )
            {
                groupUsersGroups.Add(new GroupUsersGroups
                {
                    Group = groupPrincipal,
                    UsersGroups = GetUserGroups(user)
                });
            }
            return groupUsersGroups;
        }*/

        public static IEnumerable<UserDirectReports>
            GetUsersDirectReportsFromContext(
            PrincipalContext principalContext)
        {
            return GetUsersDirectReportsFromUsers(
                GetUsers(principalContext));
        }

        public static IEnumerable<UserDirectReports>
            GetUsersDirectReportsFromUsers(
            IEnumerable<UserPrincipal> users)
        {
            return users.Select(GetUserDirectReports).ToList();
        }

        public static UserDirectReports GetUserDirectReports(
            UserPrincipal user)
        {
            return new UserDirectReports
            {
                User = user,
                DirectReports = user.GetDirectReports()
            };
        }

        public static IEnumerable<GroupPrincipal> GetGroups(
            PrincipalContext principalContext)
        {
            return GetGroups(GetUsers(principalContext));
        }

        public static IEnumerable<GroupPrincipal> GetGroups(
            IEnumerable<UserPrincipal> users)
        {
            var groups = new HashSet<GroupPrincipal>();
            foreach (var user in users)
            {
                try
                {
                    groups.UnionWith(user.GetGroups().OfType<GroupPrincipal>()
                        .ToList());
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
            }

            return groups;
        }

        public static IEnumerable<UserGroups> GetUsersGroups(
            PrincipalContext principalContext)
        {
            return GetUsersGroups(GetUsers(principalContext));
        }

        public static IEnumerable<UserGroups> GetUsersGroups(
            IEnumerable<UserPrincipal> users)
        {
            return users.Select(GetUserGroups).ToList();
        }

        public static UserGroups GetUserGroups(UserPrincipal user)
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
            return new UserGroups
            {
                User = user,
                Groups = groups
            };
        }

        public static IEnumerable<UserPrincipal> GetUsers(
            PrincipalContext context)
        {
            IEnumerable<UserPrincipal> users;
            using (var searcher = new PrincipalSearcher(
                new UserPrincipal(context)))
            {
                users = searcher.FindAll().OfType<UserPrincipal>().ToList();
            }
            return users;
        }

        public static IEnumerable<UserPrincipal> GetUsers(
            GroupPrincipal group)
        {
            return group.GetMembers().OfType<UserPrincipal>().ToList();
        }

        public static IEnumerable<UserPrincipal> GetUsersFromGroups(
            IEnumerable<GroupPrincipal> groups)
        {
            var users = new List<UserPrincipal>();
            foreach (var group in groups)
            {
                users.AddRange(GetUsers(group));
            }
            return users;
        }

        public IEnumerable<ComputerPrincipal> GetOuComputers()
        {
            return GetComputersFromContext(PrincipalContext);
        }

        public IEnumerable<UserDirectReports> GetOuUsersDirectReports()
        {
            return GetUsersDirectReportsFromContext(PrincipalContext);
        }

        public IEnumerable<GroupPrincipal> GetOuGroups()
        {
            return GetGroups(PrincipalContext);
        }

        public IEnumerable<UserPrincipal> GetOuUsers()
        {
            return GetUsers(PrincipalContext);
        }

        public IEnumerable<UserGroups> GetOuUsersGroups()
        {
            return GetUsersGroups(PrincipalContext);
        }
    }

    public interface IUserPrincipalExtender
    {
        UserPrincipal User
        {
            get;
            set;
        }
    }

    public interface IGroupPrincipalListExtender
    {
        IEnumerable<GroupPrincipal> Groups
        {
            get;
            set;
        }
    }

    public class UserDirectReports : IUserPrincipalExtender, IDisposable
    {
        public IEnumerable<UserPrincipal> DirectReports
        {
            get;
            set;
        }

        public void Dispose()
        {
            User?.Dispose();
            foreach (var directReport in DirectReports)
                directReport?.Dispose();
        }

        public UserPrincipal User
        {
            get;
            set;
        }
    }

    public class UserGroups : IUserPrincipalExtender,
        IGroupPrincipalListExtender, IDisposable
    {
        public void Dispose()
        {
            User?.Dispose();
            foreach (var group in Groups)
                group?.Dispose();
        }

        public IEnumerable<GroupPrincipal> Groups
        {
            get;
            set;
        }

        public UserPrincipal User
        {
            get;
            set;
        }
    }

    public class ComputerGroups : IGroupPrincipalListExtender, IDisposable
    {
        public ComputerPrincipal Computer
        {
            get;
            set;
        }

        public void Dispose()
        {
            Computer?.Dispose();
            foreach (var group in Groups)
                group?.Dispose();
        }

        public IEnumerable<GroupPrincipal> Groups
        {
            get;
            set;
        }
    }

    public interface IGroupPrincipalExtender
    {
        GroupPrincipal Group
        {
            get;
            set;
        }
    }

    public class GroupUsers : IGroupPrincipalExtender, IDisposable
    {
        public IEnumerable<UserPrincipal> Users
        {
            get;
            set;
        }

        public void Dispose()
        {
            Group?.Dispose();
            foreach (var user in Users)
                user?.Dispose();
        }

        public GroupPrincipal Group
        {
            get;
            set;
        }
    }

    public class GroupComputers : IGroupPrincipalExtender, IDisposable
    {
        public IEnumerable<ComputerPrincipal> Computers
        {
            get;
            set;
        }

        public void Dispose()
        {
            Group?.Dispose();
            foreach (var computer in Computers)
                computer?.Dispose();
        }

        public GroupPrincipal Group
        {
            get;
            set;
        }
    }

    public class GroupUsersGroups : IGroupPrincipalExtender, IDisposable
    {
        public GroupPrincipal Group
        {
            get;
            set;
        }

        public IEnumerable<UserGroups> UsersGroups
        {
            get;
            set;
        }

        public void Dispose()
        {
            Group?.Dispose();
            foreach(var userGroups in UsersGroups)
                userGroups?.Dispose();
        }
    }
}