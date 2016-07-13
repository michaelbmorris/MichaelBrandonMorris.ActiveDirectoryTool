using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using static ActiveDirectoryTool.QueryType;
using static ActiveDirectoryTool.SimplifiedQueryType;
using static ActiveDirectoryTool.ActiveDirectoryProperty;

namespace ActiveDirectoryTool
{
    public enum QueryType
    {
        ComputersGroups,
        ComputersSummaries,
        DirectReportsDirectReports,
        DirectReportsGroups,
        DirectReportsSummaries,
        GroupsComputers,
        GroupsUsers,
        GroupsUsersDirectReports,
        GroupsUsersGroups,
        GroupsSummaries,
        UsersDirectReports,
        UsersGroups,
        UsersSummaries,
        OuComputers,
        OuGroups,
        OuGroupsUsers,
        OuUsers,
        OuUsersDirectReports,
        OuUsersGroups
    }

    internal enum SimplifiedQueryType
    {
        DirectReports,
        Groups,
        Summaries
    }

    public class Query
    {
        private static readonly ActiveDirectoryProperty[]
            DefaultComputerProperties =
            {
                ComputerName,
                ComputerDescription,
                ComputerDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultComputerGroupsProperties =
            {
                ComputerName,
                GroupName,
                ComputerDistinguishedName,
                GroupDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultGroupProperties =
            {
                GroupName,
                GroupManagedBy,
                GroupDescription,
                GroupDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultGroupComputersProperties =
            {
                GroupName,
                ComputerName,
                GroupDistinguishedName,
                ComputerDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultGroupUsersProperties =
            {
                ContainerGroupName,
                UserName,
                ContainerGroupDistinguishedName,
                UserDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultGroupUsersDirectReportsProperties =
            {
                ContainerGroupName,
                UserName,
                DirectReportName,
                ContainerGroupDistinguishedName,
                UserDistinguishedName,
                DirectReportDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultGroupUsersGroupsProperties =
            {
                ContainerGroupName,
                UserName,
                GroupName,
                ContainerGroupDistinguishedName,
                UserDistinguishedName,
                GroupDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultUserProperties =
            {
                UserSurname,
                UserGivenName,
                UserDisplayName,
                UserSamAccountName,
                UserIsActive,
                UserIsAccountLockedOut,
                UserLastLogon,
                UserDescription,
                UserTitle,
                UserCompany,
                UserManager,
                UserHomeDrive,
                UserHomeDirectory,
                UserScriptPath,
                UserEmailAddress,
                UserStreetAddress,
                UserCity,
                UserState,
                UserVoiceTelephoneNumber,
                UserPager,
                UserMobile,
                UserFax,
                UserVoip,
                UserSip,
                UserUserPrincipalName,
                UserDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultUserDirectReportsProperties =
            {
                UserName,
                DirectReportName,
                UserDistinguishedName,
                DirectReportDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultUserGroupsProperties =
            {
                UserName,
                GroupName,
                UserDistinguishedName,
                GroupDistinguishedName
            };

        private readonly Scope _activeDirectoryScope;
        private readonly DataPreparer _dataPreparer;
        private readonly IEnumerable<string> _distinguishedNames;
        private readonly PrincipalContext _principalContext;

        private CancellationTokenSource _cancellationTokenSource;

        public Query(
            QueryType queryType,
            Scope activeDirectoryScope = null,
            IEnumerable<string> distinguishedNames = null)
        {
            QueryType = queryType;
            _principalContext = new PrincipalContext(ContextType.Domain);
            _activeDirectoryScope = activeDirectoryScope;
            _distinguishedNames = distinguishedNames;
            if (QueryTypeIsOu())
            {
                _dataPreparer = SetUpOuDataPreparer();
            }
            else if (QueryTypeIsComputer())
            {
                _dataPreparer = SetUpComputerDataPreparer();
            }
            else if (QueryTypeIsDirectReportOrUser())
            {
                _dataPreparer = SetUpDirectReportOrUserDataPreparer();
            }
            else if (QueryTypeIsGroup())
            {
                _dataPreparer = SetUpGroupDataPreparer();
            }
        }

        private CancellationToken CancellationToken
            => _cancellationTokenSource.Token;

        public IEnumerable<ExpandoObject> Data { get; private set; }

        public string Name => Scope + " - " + QueryType;

        public QueryType QueryType { get; }

        public string Scope { get; private set; }

        public void Cancel()
        {
            _cancellationTokenSource?.Cancel();
        }

        public void DisposeData()
        {
            Data = null;
        }

        public async Task Execute()
        {
            _cancellationTokenSource = new CancellationTokenSource();

            var task = Task.Run(
                () =>
                {
                    Data = GetData(_dataPreparer);
                },
                _cancellationTokenSource.Token);
            await task;
        }

        private ComputerPrincipal GetComputerPrincipal(
            string distinguishedName)
        {
            return ComputerPrincipal.FindByIdentity(
                _principalContext, distinguishedName);
        }

        private IEnumerable<ComputerPrincipal> GetComputerPrincipals()
        {
            return _distinguishedNames.Select(GetComputerPrincipal);
        }

        private static IEnumerable<ExpandoObject> GetData(
            DataPreparer dataPreparer)
        {
            return new List<ExpandoObject>(dataPreparer.GetResults());
        }

        private GroupPrincipal GetGroupPrincipal(
            string distinguishedName)
        {
            return GroupPrincipal.FindByIdentity(
                _principalContext, distinguishedName);
        }

        private IEnumerable<GroupPrincipal> GetGroupPrincipals()
        {
            return _distinguishedNames.Select(GetGroupPrincipal);
        }

        private UserPrincipal GetUserPrincipal(string distinguishedName)
        {
            return UserPrincipal.FindByIdentity(
                _principalContext, distinguishedName);
        }

        private IEnumerable<UserPrincipal> GetUserPrincipals()
        {
            return _distinguishedNames.Select(GetUserPrincipal);
        }

        private bool QueryTypeIsComputer()
        {
            return QueryType == ComputersGroups ||
                   QueryType == ComputersSummaries;
        }

        private bool QueryTypeIsDirectReportOrUser()
        {
            return QueryType == UsersDirectReports ||
                   QueryType == UsersGroups ||
                   QueryType == UsersSummaries ||
                   QueryType == DirectReportsDirectReports ||
                   QueryType == DirectReportsGroups ||
                   QueryType == DirectReportsSummaries;
        }

        private bool QueryTypeIsGroup()
        {
            return QueryType == GroupsComputers ||
                   QueryType == GroupsUsers ||
                   QueryType == GroupsUsersDirectReports ||
                   QueryType == GroupsUsersGroups ||
                   QueryType == GroupsSummaries;
        }

        private bool QueryTypeIsOu()
        {
            return QueryType == OuComputers ||
                   QueryType == OuGroups ||
                   QueryType == OuGroupsUsers ||
                   QueryType == OuUsers ||
                   QueryType == OuUsersDirectReports ||
                   QueryType == OuUsersGroups;
        }

        private DataPreparer SetUpComputerDataPreparer()
        {
            var computerPrincipals = GetComputerPrincipals();
            Scope = "Computers";
            var computerDataPreparers =
                new Dictionary<QueryType, Func<DataPreparer>>
                {
                    [ComputersGroups] = () =>
                    {
                        return new DataPreparer
                        {
                            Data = new Lazy<IEnumerable<object>>(
                                () =>
                                    Searcher.GetComputersGroups(
                                        computerPrincipals,
                                        CancellationToken)),
                            Properties = DefaultComputerGroupsProperties
                        };
                    },
                    [ComputersSummaries] = () =>
                    {
                        return new DataPreparer
                        {
                            Data = new Lazy<IEnumerable<object>>(
                                () =>
                                    computerPrincipals),
                            Properties = DefaultComputerProperties
                        };
                    }
                };
            return computerDataPreparers[QueryType]();
        }

        private DataPreparer SetUpDirectReportOrUserDataPreparer()
        {
            var userPrincipals = GetUserPrincipals();
            Scope = "Users";
            var simplifiedQueryTypes =
                new Dictionary<QueryType, SimplifiedQueryType>
                {
                    [DirectReportsDirectReports] = DirectReports,
                    [DirectReportsGroups] = Groups,
                    [DirectReportsSummaries] = Summaries,
                    [UsersDirectReports] = DirectReports,
                    [UsersGroups] = Groups,
                    [UsersSummaries] = Summaries
                };
            var directReportOrUserDataPreparers = new Dictionary
                <SimplifiedQueryType, Func<DataPreparer>>
            {
                [DirectReports] = () =>
                {
                    return new DataPreparer
                    {
                        Data = new Lazy<IEnumerable<object>>(
                            () =>
                                Searcher.GetUsersDirectReports(
                                    userPrincipals, CancellationToken)),
                        Properties = DefaultUserDirectReportsProperties
                    };
                },
                [Groups] = () =>
                {
                    return new DataPreparer
                    {
                        Data = new Lazy<IEnumerable<object>>(
                            () =>
                                Searcher.GetUsersGroups(
                                    userPrincipals, CancellationToken)),
                        Properties = DefaultUserGroupsProperties
                    };
                },
                [Summaries] = () =>
                {
                    return new DataPreparer
                    {
                        Data = new Lazy<IEnumerable<object>>(
                            () => userPrincipals),
                        Properties = DefaultUserProperties
                    };
                }
            };
            return directReportOrUserDataPreparers[
                simplifiedQueryTypes[QueryType]]();
        }

        private DataPreparer SetUpGroupDataPreparer()
        {
            var groupPrincipals = GetGroupPrincipals();
            Scope = "Groups";
            var groupDataPreparers =
                new Dictionary<QueryType, Func<DataPreparer>>
                {
                    [GroupsComputers] = () =>
                    {
                        return new DataPreparer
                        {
                            Data = new Lazy<IEnumerable<object>>(
                                () =>
                                    Searcher
                                        .GetComputerPrincipals(
                                            groupPrincipals, CancellationToken)),
                            Properties = DefaultGroupComputersProperties
                        };
                    },
                    [GroupsSummaries] = () =>
                    {
                        return new DataPreparer
                        {
                            Data = new Lazy<IEnumerable<object>>(
                                () =>
                                    groupPrincipals),
                            Properties = DefaultGroupProperties
                        };
                    },
                    [GroupsUsers] = () =>
                    {
                        return new DataPreparer
                        {
                            Data = new Lazy<IEnumerable<object>>(
                                () =>
                                    Searcher.GetGroupsUsers(
                                        groupPrincipals, CancellationToken)),
                            Properties = DefaultGroupUsersProperties
                        };
                    },
                    [GroupsUsersDirectReports] = () =>
                    {
                        return new DataPreparer
                        {
                            Data = new Lazy<IEnumerable<object>>(
                                () =>
                                    Searcher
                                        .GetGroupsUsersDirectReports(
                                            groupPrincipals, CancellationToken)),
                            Properties =
                                DefaultGroupUsersDirectReportsProperties
                        };
                    },
                    [GroupsUsersGroups] = () =>
                    {
                        return new DataPreparer
                        {
                            Data = new Lazy<IEnumerable<object>>(
                                () =>
                                    Searcher
                                        .GetGroupsUsersGroups(
                                            groupPrincipals, CancellationToken)),
                            Properties = DefaultGroupUsersGroupsProperties
                        };
                    }
                };
            return groupDataPreparers[QueryType]();
        }

        private DataPreparer SetUpOuDataPreparer()
        {
            Scope = _activeDirectoryScope.Context;
            var activeDirectorySearcher = new Searcher(
                _activeDirectoryScope);
            var ouDataPreparers = new Dictionary<QueryType, Func<DataPreparer>>
            {
                [OuComputers] = () =>
                {
                    return new DataPreparer
                    {
                        Data = new Lazy<IEnumerable<object>>(
                            () =>
                                activeDirectorySearcher.GetOuComputerPrincipals
                                    ()),
                        Properties = DefaultComputerProperties
                    };
                },
                [OuGroups] = () =>
                {
                    return new DataPreparer
                    {
                        Data = new Lazy<IEnumerable<object>>(
                            () =>
                                activeDirectorySearcher.GetOuGroupPrincipals()),
                        Properties = DefaultGroupProperties
                    };
                },
                [OuGroupsUsers] = () =>
                {
                    return new DataPreparer
                    {
                        Data = new Lazy<IEnumerable<object>>(
                            () =>
                                activeDirectorySearcher.GetOuGroupsUsers(
                                    CancellationToken)),
                        Properties = DefaultGroupUsersProperties
                    };
                },
                [OuUsers] = () =>
                {
                    return new DataPreparer
                    {
                        Data = new Lazy<IEnumerable<object>>(
                            () =>
                                activeDirectorySearcher.GetOuUserPrincipals()),
                        Properties = DefaultUserProperties
                    };
                },
                [OuUsersDirectReports] = () =>
                {
                    return new DataPreparer
                    {
                        Data = new Lazy<IEnumerable<object>>(
                            () =>
                                activeDirectorySearcher.GetOuUsersDirectReports
                                    (
                                        CancellationToken)),
                        Properties = DefaultUserDirectReportsProperties
                    };
                },
                [OuUsersGroups] = () =>
                {
                    return new DataPreparer
                    {
                        Data = new Lazy<IEnumerable<object>>(
                            () => activeDirectorySearcher.GetOuUsersGroups(
                                CancellationToken)),
                        Properties = DefaultUserGroupsProperties
                    };
                }
            };
            return ouDataPreparers[QueryType]();
        }
    }
}