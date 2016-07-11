using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using static ActiveDirectoryToolWpf.QueryType;
using static ActiveDirectoryToolWpf.SimplifiedQueryType;

namespace ActiveDirectoryToolWpf
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

    public class ActiveDirectoryQuery
    {
        private static readonly ActiveDirectoryAttribute[]
            DefaultComputerAttributes =
            {
                ActiveDirectoryAttribute.ComputerName,
                ActiveDirectoryAttribute.ComputerDescription,
                ActiveDirectoryAttribute.ComputerDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultComputerGroupsAttributes =
            {
                ActiveDirectoryAttribute.ComputerName,
                ActiveDirectoryAttribute.GroupName,
                ActiveDirectoryAttribute.ComputerDistinguishedName,
                ActiveDirectoryAttribute.GroupDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultGroupAttributes =
            {
                ActiveDirectoryAttribute.GroupName,
                ActiveDirectoryAttribute.GroupManagedBy,
                ActiveDirectoryAttribute.GroupDescription,
                ActiveDirectoryAttribute.GroupDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultGroupComputersAttributes =
            {
                ActiveDirectoryAttribute.GroupName,
                ActiveDirectoryAttribute.ComputerName,
                ActiveDirectoryAttribute.GroupDistinguishedName,
                ActiveDirectoryAttribute.ComputerDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultGroupUsersAttributes =
            {
                ActiveDirectoryAttribute.ContainerGroupName,
                ActiveDirectoryAttribute.UserName,
                ActiveDirectoryAttribute.ContainerGroupDistinguishedName,
                ActiveDirectoryAttribute.UserDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultGroupUsersDirectReportsAttributes =
            {
                ActiveDirectoryAttribute.ContainerGroupName,
                ActiveDirectoryAttribute.UserName,
                ActiveDirectoryAttribute.DirectReportName,
                ActiveDirectoryAttribute.ContainerGroupDistinguishedName,
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.DirectReportDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultGroupUsersGroupsAttributes =
            {
                ActiveDirectoryAttribute.ContainerGroupName,
                ActiveDirectoryAttribute.UserName,
                ActiveDirectoryAttribute.GroupName,
                ActiveDirectoryAttribute.ContainerGroupDistinguishedName,
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.GroupDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultUserAttributes =
            {
                ActiveDirectoryAttribute.UserSurname,
                ActiveDirectoryAttribute.UserGivenName,
                ActiveDirectoryAttribute.UserDisplayName,
                ActiveDirectoryAttribute.UserSamAccountName,
                ActiveDirectoryAttribute.UserIsActive,
                ActiveDirectoryAttribute.UserIsAccountLockedOut,
                ActiveDirectoryAttribute.UserLastLogon,
                ActiveDirectoryAttribute.UserDescription,
                ActiveDirectoryAttribute.UserTitle,
                ActiveDirectoryAttribute.UserCompany,
                ActiveDirectoryAttribute.UserManager,
                ActiveDirectoryAttribute.UserHomeDrive,
                ActiveDirectoryAttribute.UserHomeDirectory,
                ActiveDirectoryAttribute.UserScriptPath,
                ActiveDirectoryAttribute.UserEmailAddress,
                ActiveDirectoryAttribute.UserStreetAddress,
                ActiveDirectoryAttribute.UserCity,
                ActiveDirectoryAttribute.UserState,
                ActiveDirectoryAttribute.UserVoiceTelephoneNumber,
                ActiveDirectoryAttribute.UserPager,
                ActiveDirectoryAttribute.UserMobile,
                ActiveDirectoryAttribute.UserFax,
                ActiveDirectoryAttribute.UserVoip,
                ActiveDirectoryAttribute.UserSip,
                ActiveDirectoryAttribute.UserUserPrincipalName,
                ActiveDirectoryAttribute.UserDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultUserDirectReportsAttributes =
            {
                ActiveDirectoryAttribute.UserName,
                ActiveDirectoryAttribute.DirectReportName,
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.DirectReportDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultUserGroupsAttributes =
            {
                ActiveDirectoryAttribute.UserName,
                ActiveDirectoryAttribute.GroupName,
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.GroupDistinguishedName
            };

        private readonly ActiveDirectoryScope _activeDirectoryScope;
        private readonly DataPreparer _dataPreparer;

        private readonly IEnumerable<string> _distinguishedNames;
        private readonly PrincipalContext _principalContext;

        private CancellationTokenSource _cancellationTokenSource;

        public ActiveDirectoryQuery(
            QueryType queryType,
            ActiveDirectoryScope activeDirectoryScope = null,
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
                () => { Data = GetData(_dataPreparer); },
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
            return _distinguishedNames.Select(GetComputerPrincipal).ToList();
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
            return _distinguishedNames.Select(GetGroupPrincipal).ToList();
        }

        private UserPrincipal GetUserPrincipal(string distinguishedName)
        {
            return UserPrincipal.FindByIdentity(
                _principalContext, distinguishedName);
        }

        private IEnumerable<UserPrincipal> GetUserPrincipals()
        {
            return _distinguishedNames.Select(GetUserPrincipal).ToList();
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
                                    ActiveDirectorySearcher.GetComputersGroups(
                                        computerPrincipals, CancellationToken)),
                            Attributes = DefaultComputerGroupsAttributes
                        };
                    },
                    [ComputersSummaries] = () =>
                    {
                        return new DataPreparer
                        {
                            Data = new Lazy<IEnumerable<object>>(
                                () =>
                                    computerPrincipals),
                            Attributes = DefaultComputerAttributes
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
                                ActiveDirectorySearcher.GetUsersDirectReports(
                                    userPrincipals, CancellationToken)),
                        Attributes = DefaultUserDirectReportsAttributes
                    };
                },
                [Groups] = () =>
                {
                    return new DataPreparer
                    {
                        Data = new Lazy<IEnumerable<object>>(
                            () =>
                                ActiveDirectorySearcher.GetUsersGroups(
                                    userPrincipals, CancellationToken)),
                        Attributes = DefaultUserGroupsAttributes
                    };
                },
                [Summaries] = () =>
                {
                    return new DataPreparer
                    {
                        Data = new Lazy<IEnumerable<object>>(
                            () =>
                                userPrincipals),
                        Attributes = DefaultUserAttributes
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
                                    ActiveDirectorySearcher
                                        .GetComputerPrincipals(
                                            groupPrincipals, CancellationToken)),
                            Attributes = DefaultGroupComputersAttributes
                        };
                    },
                    [GroupsSummaries] = () =>
                    {
                        return new DataPreparer
                        {
                            Data = new Lazy<IEnumerable<object>>(
                                () =>
                                    groupPrincipals),
                            Attributes = DefaultGroupAttributes
                        };
                    },
                    [GroupsUsers] = () =>
                    {
                        return new DataPreparer
                        {
                            Data = new Lazy<IEnumerable<object>>(
                                () =>
                                    ActiveDirectorySearcher.GetGroupsUsers(
                                        groupPrincipals, CancellationToken)),
                            Attributes = DefaultGroupUsersAttributes
                        };
                    },
                    [GroupsUsersDirectReports] = () =>
                    {
                        return new DataPreparer
                        {
                            Data = new Lazy<IEnumerable<object>>(
                                () =>
                                    ActiveDirectorySearcher
                                        .GetGroupsUsersDirectReports(
                                            groupPrincipals, CancellationToken)),
                            Attributes =
                                DefaultGroupUsersDirectReportsAttributes
                        };
                    },
                    [GroupsUsersGroups] = () =>
                    {
                        return new DataPreparer
                        {
                            Data = new Lazy<IEnumerable<object>>(
                                () =>
                                    ActiveDirectorySearcher
                                        .GetGroupsUsersGroups(
                                            groupPrincipals, CancellationToken)),
                            Attributes = DefaultGroupUsersGroupsAttributes
                        };
                    }
                };
            return groupDataPreparers[QueryType]();
        }

        private DataPreparer SetUpOuDataPreparer()
        {
            Scope = _activeDirectoryScope.Context;
            var activeDirectorySearcher = new ActiveDirectorySearcher(
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
                        Attributes = DefaultComputerAttributes
                    };
                },
                [OuGroups] = () =>
                {
                    return new DataPreparer
                    {
                        Data = new Lazy<IEnumerable<object>>(
                            () =>
                                activeDirectorySearcher.GetOuGroupPrincipals()),
                        Attributes = DefaultGroupAttributes
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
                        Attributes = DefaultGroupUsersAttributes
                    };
                },
                [OuUsers] = () =>
                {
                    return new DataPreparer
                    {
                        Data = new Lazy<IEnumerable<object>>(
                            () =>
                                activeDirectorySearcher.GetOuUserPrincipals()),
                        Attributes = DefaultUserAttributes
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
                        Attributes = DefaultUserDirectReportsAttributes
                    };
                },
                [OuUsersGroups] = () =>
                {
                    return new DataPreparer
                    {
                        Data = new Lazy<IEnumerable<object>>(
                            () =>
                                activeDirectorySearcher.GetOuUsersGroups(
                                    CancellationToken)),
                        Attributes = DefaultUserGroupsAttributes
                    };
                }
            };
            return ouDataPreparers[QueryType]();
        }
    }
}