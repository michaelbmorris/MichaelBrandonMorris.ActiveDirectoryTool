using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.Linq;

namespace ActiveDirectoryToolWpf
{
    public interface IComputerPrincipal
    {
        ComputerPrincipal Computer
        {
            get;
            set;
        }
    }

    public interface IComputers
    {
        IEnumerable<ComputerPrincipal> Computers
        {
            get;
            set;
        }
    }

    public interface IDirectReports
    {
        IEnumerable<UserPrincipal> DirectReports
        {
            get;
            set;
        }
    }

    public interface IGroup
    {
        GroupPrincipal Group
        {
            get;
            set;
        }
    }

    public interface IGroups
    {
        IEnumerable<GroupPrincipal> Groups
        {
            get;
            set;
        }
    }

    public interface IUser
    {
        UserPrincipal User
        {
            get;
            set;
        }
    }

    public interface IUserGroups
    {
        IEnumerable<UserGroups> UsersGroups
        {
            get;
            set;
        }
    }

    public interface IUsers
    {
        IEnumerable<UserPrincipal> Users
        {
            get;
            set;
        }
    }

    public interface IUsersDirectReports
    {
        IEnumerable<UserDirectReports> UsersDirectReports
        {
            get;
            set;
        }
    }

    public interface IUsersGroups
    {
        IEnumerable<UserGroups> UsersGroups
        {
            get;
            set;
        }
    }

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

        public static GroupComputers GetComputers(GroupPrincipal groupPrincipal)
        {
            return new GroupComputers
            {
                Group = groupPrincipal,
                Computers = groupPrincipal.Members.OfType<ComputerPrincipal>()
            };
        }

        public static IEnumerable<ComputerPrincipal> GetComputers(
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

        public static GroupUsersDirectReports GetGroupUsersDirectReports(
            GroupPrincipal groupPrincipal)
        {
            return new GroupUsersDirectReports
            {
                Group = groupPrincipal,
                UsersDirectReports = GetUsersDirectReports(groupPrincipal)
            };
        }

        public static GroupUsersGroups GetGroupUsersGroups(
            GroupPrincipal groupPrincipal)
        {
            return new GroupUsersGroups
            {
                Group = groupPrincipal,
                UsersGroups = GetUsersGroups(groupPrincipal)
            };
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

        public static GroupUsers GetUsers(GroupPrincipal groupPrincipal)
        {
            return new GroupUsers
            {
                Group = groupPrincipal,
                Users = groupPrincipal.Members.OfType<UserPrincipal>()
            };
        }

        public static IEnumerable<GroupUsers> GetUsers(
            IEnumerable<GroupPrincipal> groupPrincipals)
        {
            return groupPrincipals.Select(GetUsers).ToList();
        }

        public static IEnumerable<UserDirectReports> GetUsersDirectReports(
            GroupPrincipal groupPrincipal)
        {
            return groupPrincipal.Members.OfType<UserPrincipal>().Select(
                GetUserDirectReports).ToList();
        }

        public static IEnumerable<UserDirectReports>
            GetUsersDirectReports(PrincipalContext principalContext)
        {
            return GetUsersDirectReports(GetUsers(principalContext));
        }

        public static IEnumerable<UserDirectReports> GetUsersDirectReports(
            IEnumerable<UserPrincipal> userPrincipals)
        {
            return userPrincipals.Select(GetUserDirectReports).ToList();
        }

        public static IEnumerable<UserGroups> GetUsersGroups(
            GroupPrincipal groupPrincipal)
        {
            return groupPrincipal.Members.OfType<UserPrincipal>().Select(
                GetUserGroups).ToList();
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

        public IEnumerable<ComputerPrincipal> GetComputers()
        {
            return GetComputers(PrincipalContext);
        }

        public IEnumerable<ComputerGroups> GetComputersGroups()
        {
            return GetComputersGroups(PrincipalContext);
        }

        public IEnumerable<GroupPrincipal> GetGroups()
        {
            return GetGroups(PrincipalContext);
        }

        public IEnumerable<UserPrincipal> GetUsers()
        {
            return GetUsers(PrincipalContext);
        }

        public IEnumerable<UserDirectReports> GetUsersDirectReports()
        {
            return GetUsersDirectReports(PrincipalContext);
        }

        public IEnumerable<UserGroups> GetUsersGroups()
        {
            return GetUsersGroups(PrincipalContext);
        }
    }

    public class ComputerGroups : IComputerPrincipal, IDisposable, IGroups
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

    public class GroupComputers : IComputers, IDisposable, IGroup
    {
        public IEnumerable<ComputerPrincipal> Computers
        {
            get;
            set;
        }

        public void Dispose()
        {
            Group?.Dispose();
            if (Computers == null) return;
            foreach (var computer in Computers)
                computer?.Dispose();
        }

        public GroupPrincipal Group
        {
            get;
            set;
        }
    }

    public class GroupUsers : IGroup, IDisposable, IUsers
    {
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

        public IEnumerable<UserPrincipal> Users
        {
            get;
            set;
        }
    }

    public class GroupUsersDirectReports : IDisposable, IGroup,
        IUsersDirectReports
    {
        public void Dispose()
        {
            Group?.Dispose();
            foreach (var userDirectReports in UsersDirectReports)
                userDirectReports?.Dispose();
        }

        public GroupPrincipal Group
        {
            get;
            set;
        }

        public IEnumerable<UserDirectReports> UsersDirectReports
        {
            get;
            set;
        }
    }

    public class GroupUsersGroups : IDisposable, IGroup, IUsersGroups
    {
        public void Dispose()
        {
            Group?.Dispose();
            foreach (var userGroups in UsersGroups)
                userGroups?.Dispose();
        }

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
    }

    public class UserDirectReports : IDirectReports, IDisposable, IUser
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

    public class UserGroups : IDisposable, IGroups, IUser
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
}