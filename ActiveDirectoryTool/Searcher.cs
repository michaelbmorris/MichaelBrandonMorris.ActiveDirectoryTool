using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Threading;
using Extensions.PrincipalExtensions;

namespace ActiveDirectoryTool
{
    public interface IComputer
    {
        ComputerPrincipal Computer { get; set; }
    }

    public interface IComputers
    {
        IEnumerable<ComputerPrincipal> Computers { get; set; }
    }

    public interface IDirectReports
    {
        IEnumerable<UserPrincipal> DirectReports { get; set; }
    }

    public interface IGroup
    {
        GroupPrincipal Group { get; set; }
    }

    public interface IGroups
    {
        IEnumerable<GroupPrincipal> Groups { get; set; }
    }

    public interface IUser
    {
        UserPrincipal User { get; set; }
    }

    public interface IUserGroups
    {
        IEnumerable<UserGroups> UsersGroups { get; set; }
    }

    public interface IUsers
    {
        IEnumerable<UserPrincipal> Users { get; set; }
    }

    public interface IUsersDirectReports
    {
        IEnumerable<UserDirectReports> UsersDirectReports { get; set; }
    }

    public interface IUsersGroups
    {
        IEnumerable<UserGroups> UsersGroups { get; set; }
    }

    public static class Searcher
    {
        private const char Asterix = '*';

        public static ComputerGroups GetComputerGroups(
            ComputerPrincipal computerPrincipal,
            CancellationToken cancellationToken)
        {
            return new ComputerGroups
            {
                Computer = computerPrincipal,
                Groups = computerPrincipal.GetGroupPrincipals()
            };
        }

        public static IEnumerable<ComputerPrincipal> GetComputerPrincipals(
            PrincipalContext principalContext,
            CancellationToken cancellationToken)
        {
            IEnumerable<ComputerPrincipal> computers;
            using (var computerPrincipal =
                new ComputerPrincipal(principalContext))
            {
                using (var searcher = new PrincipalSearcher(computerPrincipal))
                {
                    computers = searcher.GetAllComputerPrincipals();
                }
            }
            
            return computers;
        }

        public static IEnumerable<ComputerPrincipal> GetComputerPrincipals(
            GroupPrincipal groupPrincipal,
            CancellationToken cancellationToken)
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
                groupsComputers.Add(
                    GetGroupComputers(groupPrincipal, cancellationToken));
            }
            return groupsComputers;
        }

        public static IEnumerable<ComputerGroups> GetComputersGroups(
            PrincipalContext principalContext,
            CancellationToken cancellationToken)
        {
            var computersGroups = new List<ComputerGroups>();
            using (var searchPrincipal =
                new ComputerPrincipal(principalContext))
            {
                using (var searcher = new PrincipalSearcher(searchPrincipal))
                {
                    foreach (var computerPrincipal in searcher
                        .GetAllComputerPrincipals())
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        computersGroups.Add(
                            GetComputerGroups(
                                computerPrincipal, cancellationToken));
                    }
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
                computersGroups.Add(
                    GetComputerGroups(computerPrincipal, cancellationToken));
            }
            return computersGroups;
        }

        public static GroupComputers GetGroupComputers(
            GroupPrincipal groupPrincipal,
            CancellationToken cancellationToken)
        {
            return new GroupComputers
            {
                Group = groupPrincipal,
                Computers = GetComputerPrincipals(
                    groupPrincipal, cancellationToken)
            };
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
            PrincipalContext principalContext,
            CancellationToken cancellationToken)
        {
            IEnumerable<GroupPrincipal> groupPrincipals;
            using (var searchPrincipal = new GroupPrincipal(principalContext))
            {
                using (var searcher = new PrincipalSearcher(searchPrincipal))
                {
                    groupPrincipals = searcher.GetAllGroupPrincipals();
                }
            }
            
            return groupPrincipals;
        }

        public static IEnumerable<GroupUsers> GetGroupsUsers(
            PrincipalContext principalContext,
            CancellationToken cancellationToken)
        {
            return GetGroupsUsers(
                GetGroupPrincipals(
                    principalContext, cancellationToken), cancellationToken);
        }

        public static IEnumerable<GroupUsers> GetGroupsUsers(
            IEnumerable<GroupPrincipal> groupPrincipals,
            CancellationToken cancellationToken)
        {
            var groupsUsers = new List<GroupUsers>();
            foreach (var groupPrincipal in groupPrincipals)
            {
                cancellationToken.ThrowIfCancellationRequested();
                groupsUsers.Add(
                    new GroupUsers
                    {
                        Group = groupPrincipal,
                        Users = GetUserPrincipals(groupPrincipal)
                    });
            }
            return groupsUsers;
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

        public static UserDirectReports GetUserDirectReports(
            UserPrincipal userPrincipal)
        {
            return new UserDirectReports
            {
                User = userPrincipal,
                //DirectReports = userPrincipal.GetDirectReportUserPrincipals()
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

        public static IEnumerable<User> GetUsers(
            PrincipalContext principalContext,
            CancellationToken cancellationToken)
        {
            var stopwatch = Stopwatch.StartNew();
            var users = new List<User>();
            using (var up = new UserPrincipal(principalContext))
            {
                using (var ps = new PrincipalSearcher(up))
                {
                    using (var psr = ps.FindAll())
                    {
                        users.AddRange(
                            from UserPrincipal u in psr select new User(u));
                    }
                }
            }

            stopwatch.Stop();
            Debug.WriteLine(stopwatch.ElapsedMilliseconds / 1000);
            return users;
        }

        public static IEnumerable<UserPrincipal> GetUserPrincipals(
            PrincipalContext principalContext,
            CancellationToken cancellationToken)
        {
            var stopwatch = Stopwatch.StartNew();
            IEnumerable<UserPrincipal> userPrincipals;
            using (var searchPrincipal = new UserPrincipal(principalContext))
            {
                using (var searcher = new PrincipalSearcher(searchPrincipal))
                {
                    userPrincipals = searcher.GetAllUserPrincipals();
                }
            }

            stopwatch.Stop();
            Debug.WriteLine(stopwatch.ElapsedMilliseconds / 1000);
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
                GetUserPrincipals(
                    principalContext, cancellationToken), cancellationToken);
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
                GetUserPrincipals(
                    principalContext, cancellationToken), cancellationToken);
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

        public static IEnumerable<UserPrincipal> GetUserPrincipals(
            PrincipalContext principalContext, string searchText)
        {
            var userPrincipals = new List<UserPrincipal>();
            using (var searchPrincipal = new UserPrincipal(principalContext)
            {
                Name = Asterix + searchText + Asterix
            })
            {
                using (var principalSearcher = new PrincipalSearcher(
                    searchPrincipal))
                {
                    userPrincipals.AddRange(
                        principalSearcher.GetAllUserPrincipals());
                }
            }
            return userPrincipals;
        }

        public static IEnumerable<ComputerPrincipal> GetComputerPrincipals(
            PrincipalContext principalContext, string searchText)
        {
            var computerPrincipals = new List<ComputerPrincipal>();
            using (var searchPrincipal = new ComputerPrincipal(principalContext)
            {
                Name = Asterix + searchText + Asterix
            })
            {
                using (var principalSearcher = new PrincipalSearcher(
                    searchPrincipal))
                {
                    computerPrincipals.AddRange(
                        principalSearcher.GetAllComputerPrincipals());
                }
            }

            return computerPrincipals;
        }

        public static IEnumerable<GroupPrincipal> GetGroupPrincipals(
            PrincipalContext principalContext, string searchText)
        {
            var groupPrincipals = new List<GroupPrincipal>();
            using (var searchPrincipal = new GroupPrincipal(principalContext)
            {
                Name = Asterix + searchText + Asterix
            })
            {
                using (var principalSearcher = new PrincipalSearcher(
                    searchPrincipal))
                {
                    groupPrincipals.AddRange(
                        principalSearcher.GetAllGroupPrincipals());
                }
            }
            return groupPrincipals;
        }
    }

    public class ComputerGroups : IComputer, IGroups
    {
        public ComputerPrincipal Computer { get; set; }

        public IEnumerable<GroupPrincipal> Groups { get; set; }
    }

    public class GroupComputers : IComputers, IGroup
    {
        public IEnumerable<ComputerPrincipal> Computers { get; set; }

        public GroupPrincipal Group { get; set; }
    }

    public class GroupUsers : IGroup, IUsers
    {
        public GroupPrincipal Group { get; set; }

        public IEnumerable<UserPrincipal> Users { get; set; }
    }

    public class GroupUsersDirectReports : IGroup,
        IUsersDirectReports
    {
        public GroupPrincipal Group { get; set; }

        public IEnumerable<UserDirectReports> UsersDirectReports { get; set; }
    }

    public class GroupUsersGroups : IGroup, IUsersGroups
    {
        public GroupPrincipal Group { get; set; }

        public IEnumerable<UserGroups> UsersGroups { get; set; }
    }

    public class UserDirectReports : IDirectReports, IUser
    {
        public IEnumerable<UserPrincipal> DirectReports { get; set; }

        public UserPrincipal User { get; set; }
    }

    public class UserGroups : IGroups, IUser
    {
        public IEnumerable<GroupPrincipal> Groups { get; set; }

        public UserPrincipal User { get; set; }
    }
}