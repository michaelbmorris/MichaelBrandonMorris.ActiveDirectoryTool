using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.Threading;

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
        public ActiveDirectorySearcher(
            ActiveDirectoryScope activeDirectoryScope)
        {
            ActiveDirectoryScope = activeDirectoryScope;
        }

        private ActiveDirectoryScope ActiveDirectoryScope
        {
            get;
        }

        private PrincipalContext PrincipalContext => new PrincipalContext(
            ContextType.Domain,
            ActiveDirectoryScope.Domain,
            ActiveDirectoryScope.Context);

        public static ComputerGroups GetComputerGroups(
            ComputerPrincipal computerPrincipal)
        {
            return new ComputerGroups
            {
                Computer = computerPrincipal,
                Groups = computerPrincipal.GetGroupPrincipals()
            };
        }

        public static IEnumerable<ComputerPrincipal> GetComputerPrincipals(
            PrincipalContext principalContext)
        {
            IEnumerable<ComputerPrincipal> computers;
            using (var searcher = new PrincipalSearcher(
                new ComputerPrincipal(principalContext)))
            {
                computers = searcher.GetAllComputerPrincipals();
            }
            return computers;
        }

        public static IEnumerable<ComputerPrincipal> GetComputerPrincipals(
            GroupPrincipal groupPrincipal)
        {
            return groupPrincipal.GetComputerPrincipals();
        }

        public static IEnumerable<GroupComputers> GetComputerPrincipals(
            IEnumerable<GroupPrincipal> groupPrincipals,
            CancellationToken cancellationToken)
        {
            var groupsComputers = new List<GroupComputers>();
            foreach (var groupPrincipal in groupPrincipals)
            {
                cancellationToken.ThrowIfCancellationRequested();
                groupsComputers.Add(GetGroupComputers(groupPrincipal));
            }
            return groupsComputers;
        }

        public static GroupComputers GetGroupComputers(
            GroupPrincipal groupPrincipal)
        {
            return new GroupComputers
            {
                Group = groupPrincipal,
                Computers = GetComputerPrincipals(groupPrincipal)
            };
        }

        public static IEnumerable<ComputerGroups> GetComputersGroups(
            PrincipalContext principalContext,
            CancellationToken cancellationToken)
        {
            var computersGroups = new List<ComputerGroups>();
            using (var searcher = new PrincipalSearcher(
                new ComputerPrincipal(principalContext)))
            {
                foreach (var computerPrincipal in searcher
                    .GetAllComputerPrincipals())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    computersGroups.Add(GetComputerGroups(computerPrincipal));
                }
            }
            return computersGroups;
        }

        public static IEnumerable<ComputerGroups> GetComputersGroups(
            IEnumerable<ComputerPrincipal> computerPrincipals,
            CancellationToken cancellationToken)
        {
            var computersGroups = new List<ComputerGroups>();
            foreach (var computerPrincipal in computerPrincipals)
            {
                cancellationToken.ThrowIfCancellationRequested();
                computersGroups.Add(GetComputerGroups(computerPrincipal));
            }
            return computersGroups;
        }

        public static IEnumerable<GroupPrincipal> GetGroupPrincipals(
            IEnumerable<UserPrincipal> userPrincipals,
            CancellationToken cancellationToken)
        {
            var groupPrincipals = new HashSet<GroupPrincipal>();
            foreach (var userPrincipal in userPrincipals)
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    groupPrincipals.UnionWith(
                        userPrincipal.GetGroupPrincipals());
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
            }

            return groupPrincipals;
        }

        public static IEnumerable<GroupPrincipal> GetGroupPrincipals(
            PrincipalContext principalContext)
        {
            IEnumerable<GroupPrincipal> groupPrincipals;
            using (var searcher = new PrincipalSearcher(
                new GroupPrincipal(principalContext)))
            {
                groupPrincipals = searcher.GetAllGroupPrincipals();
            }
            return groupPrincipals;
        }

        public static IEnumerable<GroupUsers> GetGroupsUsers(
            PrincipalContext principalContext,
            CancellationToken cancellationToken)
        {
            return GetGroupsUsers(
                GetGroupPrincipals(principalContext), cancellationToken);
        }

        public static IEnumerable<GroupUsers> GetGroupsUsers(
            IEnumerable<GroupPrincipal> groupPrincipals,
            CancellationToken cancellationToken)
        {
            var groupsUsers = new List<GroupUsers>();
            foreach (var groupPrincipal in groupPrincipals)
            {
                cancellationToken.ThrowIfCancellationRequested();
                groupsUsers.Add(new GroupUsers
                {
                    Group = groupPrincipal,
                    Users = GetUserPrincipals(groupPrincipal)
                });
            }
            return groupsUsers;
        }

        public static GroupUsersDirectReports GetGroupUsersDirectReports(
            GroupPrincipal groupPrincipal,
            CancellationToken cancellationToken)
        {
            return new GroupUsersDirectReports
            {
                Group = groupPrincipal,
                UsersDirectReports = GetUsersDirectReports(
                    groupPrincipal, cancellationToken)
            };
        }

        public static IEnumerable<GroupUsersDirectReports>
            GetGroupsUsersDirectReports(
            IEnumerable<GroupPrincipal> groupPrincipals,
            CancellationToken cancellationToken)
        {
            var groupsUsersDirectReports = new List<GroupUsersDirectReports>();
            foreach (var groupPrincipal in groupPrincipals)
            {
                cancellationToken.ThrowIfCancellationRequested();
                groupsUsersDirectReports.Add(
                    GetGroupUsersDirectReports(
                        groupPrincipal, cancellationToken));
            }
            return groupsUsersDirectReports;
        }

        public static GroupUsersGroups GetGroupUsersGroups(
            GroupPrincipal groupPrincipal,
            CancellationToken cancellationToken)
        {
            return new GroupUsersGroups
            {
                Group = groupPrincipal,
                UsersGroups = GetUsersGroups(groupPrincipal, cancellationToken)
            };
        }

        public static IEnumerable<GroupUsersGroups> GetGroupsUsersGroups(
            IEnumerable<GroupPrincipal> groupPrincipals,
            CancellationToken cancellationToken)
        {
            var groupsUsersGroups = new List<GroupUsersGroups>();
            foreach (var groupPrincipal in groupPrincipals)
            {
                cancellationToken.ThrowIfCancellationRequested();
                groupsUsersGroups.Add(
                    GetGroupUsersGroups(groupPrincipal, cancellationToken));
            }
            return groupsUsersGroups;
        }

        public static UserDirectReports GetUserDirectReports(
            UserPrincipal userPrincipal)
        {
            return new UserDirectReports
            {
                User = userPrincipal,
                DirectReports = userPrincipal.GetDirectReportUserPrincipals()
            };
        }

        public static UserGroups GetUserGroups(UserPrincipal userPrincipal)
        {
            var groups = new List<GroupPrincipal>();
            try
            {
                groups.AddRange(userPrincipal.GetGroupPrincipals());
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
            return new UserGroups
            {
                User = userPrincipal,
                Groups = groups
            };
        }

        public static IEnumerable<UserPrincipal> GetUserPrincipals(
            PrincipalContext principalContext)
        {
            IEnumerable<UserPrincipal> userPrincipals;
            using (var searcher = new PrincipalSearcher(
                new UserPrincipal(principalContext)))
            {
                userPrincipals = searcher.GetAllUserPrincipals();
            }
            return userPrincipals;
        }

        public static IEnumerable<UserPrincipal> GetUserPrincipals(
            GroupPrincipal groupPrincipal)
        {
            return groupPrincipal.GetUserPrincipals();
        }

        public static IEnumerable<UserPrincipal> GetUserPrincipals(
            IEnumerable<GroupPrincipal> groupPrincipals,
            CancellationToken cancellationToken)
        {
            var userPrincipals = new List<UserPrincipal>();
            foreach (var groupPrincipal in groupPrincipals)
            {
                cancellationToken.ThrowIfCancellationRequested();
                userPrincipals.AddRange(GetUserPrincipals(groupPrincipal));
            }
            return userPrincipals;
        }

        public static IEnumerable<UserDirectReports> GetUsersDirectReports(
            GroupPrincipal groupPrincipal,
            CancellationToken cancellationToken)
        {
            var usersDirectReports = new List<UserDirectReports>();
            foreach (var userPrincipal in groupPrincipal.GetUserPrincipals())
            {
                cancellationToken.ThrowIfCancellationRequested();
                usersDirectReports.Add(GetUserDirectReports(userPrincipal));
            }
            return usersDirectReports;
        }

        public static IEnumerable<UserDirectReports> GetUsersDirectReports(
            PrincipalContext principalContext,
            CancellationToken cancellationToken)
        {
            return GetUsersDirectReports(
                GetUserPrincipals(principalContext), cancellationToken);
        }

        public static IEnumerable<UserDirectReports> GetUsersDirectReports(
            IEnumerable<UserPrincipal> userPrincipals,
            CancellationToken cancellationToken)
        {
            var usersDirectReports = new List<UserDirectReports>();
            foreach (var userPrincipal in userPrincipals)
            {
                cancellationToken.ThrowIfCancellationRequested();
                usersDirectReports.Add(GetUserDirectReports(userPrincipal));
            }
            return usersDirectReports;
        }

        public static IEnumerable<UserGroups> GetUsersGroups(
            GroupPrincipal groupPrincipal,
            CancellationToken cancellationToken)
        {
            var usersGroups = new List<UserGroups>();
            foreach (var userPrincipal in groupPrincipal.GetUserPrincipals())
            {
                cancellationToken.ThrowIfCancellationRequested();
                usersGroups.Add(GetUserGroups(userPrincipal));
            }
            return usersGroups;
        }

        public static IEnumerable<UserGroups> GetUsersGroups(
            PrincipalContext principalContext,
            CancellationToken cancellationToken)
        {
            return GetUsersGroups(
                GetUserPrincipals(principalContext), cancellationToken);
        }

        public static IEnumerable<UserGroups> GetUsersGroups(
            IEnumerable<UserPrincipal> userPrincipals,
            CancellationToken cancellationToken)
        {
            var usersGroups = new List<UserGroups>();
            foreach (var userPrincipal in userPrincipals)
            {
                cancellationToken.ThrowIfCancellationRequested();
                usersGroups.Add(GetUserGroups(userPrincipal));
            }
            return usersGroups;
        }

        public IEnumerable<ComputerPrincipal> GetOuComputerPrincipals()
        {
            return GetComputerPrincipals(PrincipalContext);
        }

        public IEnumerable<GroupPrincipal> GetOuGroupPrincipals()
        {
            return GetGroupPrincipals(PrincipalContext);
        }

        public IEnumerable<GroupUsers> GetOuGroupsUsers(
            CancellationToken cancellationToken)
        {
            return GetGroupsUsers(PrincipalContext, cancellationToken);
        }

        public IEnumerable<UserPrincipal> GetOuUserPrincipals()
        {
            return GetUserPrincipals(PrincipalContext);
        }

        public IEnumerable<UserDirectReports> GetOuUsersDirectReports(
            CancellationToken cancellationToken)
        {
            return GetUsersDirectReports(PrincipalContext, cancellationToken);
        }

        public IEnumerable<UserGroups> GetOuUsersGroups(
            CancellationToken cancellationToken)
        {
            return GetUsersGroups(PrincipalContext, cancellationToken);
        }
    }

    public class ComputerGroups : IComputerPrincipal, IGroups
    {
        public ComputerPrincipal Computer
        {
            get;
            set;
        }

        public IEnumerable<GroupPrincipal> Groups
        {
            get;
            set;
        }
    }

    public class GroupComputers : IComputers, IGroup
    {
        public IEnumerable<ComputerPrincipal> Computers
        {
            get;
            set;
        }

        public GroupPrincipal Group
        {
            get;
            set;
        }
    }

    public class GroupUsers : IGroup, IUsers
    {
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

    public class GroupUsersDirectReports : IGroup,
        IUsersDirectReports
    {
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

    public class GroupUsersGroups : IGroup, IUsersGroups
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
    }

    public class UserDirectReports : IDirectReports, IUser
    {
        public IEnumerable<UserPrincipal> DirectReports
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

    public class UserGroups : IGroups, IUser
    {
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