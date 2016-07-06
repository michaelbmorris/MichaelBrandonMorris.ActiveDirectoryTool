using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Threading;
using System.Threading.Tasks;
using static ActiveDirectoryToolWpf.QueryType;
using static ActiveDirectoryToolWpf.SimplifiedQueryType;

namespace ActiveDirectoryToolWpf
{
    public enum QueryType
    {
        ContextComputerGroups,
        ContextComputerSummary,
        ContextDirectReportDirectReports,
        ContextDirectReportGroups,
        ContextDirectReportSummary,
        ContextGroupComputers,
        ContextGroupUsers,
        ContextGroupUsersDirectReports,
        ContextGroupUsersGroups,
        ContextGroupSummary,
        ContextUserDirectReports,
        ContextUserGroups,
        ContextUserSummary,
        OuComputers,
        OuGroups,
        OuUsers,
        OuUsersDirectReports,
        OuUsersGroups
    }

    internal enum SimplifiedQueryType
    {
        DirectReports,
        Groups,
        Summary
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
                ActiveDirectoryAttribute.UserSurname,
                ActiveDirectoryAttribute.UserGivenName,
                ActiveDirectoryAttribute.UserDisplayName,
                ActiveDirectoryAttribute.UserSamAccountName,
                ActiveDirectoryAttribute.UserIsActive,
                ActiveDirectoryAttribute.UserIsAccountLockedOut,
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
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.ContainerGroupDistinguishedName
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
        private readonly PrincipalContext _principalContext;

        private readonly string _selectedItemDistinguishedName;

        private CancellationTokenSource _cancellationTokenSource;

        public ActiveDirectoryQuery(
            QueryType queryType,
            ActiveDirectoryScope activeDirectoryScope = null,
            string selectedItemDistinguishedName = null)
        {
            QueryType = queryType;
            _principalContext = new PrincipalContext(ContextType.Domain);
            _activeDirectoryScope = activeDirectoryScope;
            _selectedItemDistinguishedName = selectedItemDistinguishedName;
        }

        public IEnumerable<ExpandoObject> Data
        {
            get;
            private set;
        }

        public QueryType QueryType
        {
            get;
        }

        public string Scope
        {
            get;
            private set;
        }

        public string Name => Scope + QueryType;

        private CancellationToken CancellationToken
            => _cancellationTokenSource.Token;

        public void Cancel()
        {
            _cancellationTokenSource?.Cancel();
        }

        public async Task Execute()
        {
            _cancellationTokenSource = new CancellationTokenSource();
            DataPreparer dataPreparer = null;
            var task = Task.Run(() =>
            {
                if (QueryTypeIsOu())
                {
                    dataPreparer = SetUpOuDataPreparer();
                }
                else if (QueryTypeIsContextComputer())
                {
                    dataPreparer = SetUpComputerDataPreparer();
                }
                else if (QueryTypeIsContextDirectReportOrUser())
                {
                    dataPreparer = SetUpDirectReportOrUserDataPreparer();
                }
                else if (QueryTypeIsContextGroup())
                {
                    dataPreparer = SetUpGroupDataPreparer();
                }
                Data = GetData(dataPreparer);
            },
                _cancellationTokenSource.Token);
            await task;
        }

        private static IEnumerable<ExpandoObject> GetData(
            DataPreparer dataPreparer)
        {
            return new List<ExpandoObject>(dataPreparer.GetResults());
        }

        private ComputerPrincipal GetSelectedComputerPrincipal()
        {
            return ComputerPrincipal.FindByIdentity(
                _principalContext, _selectedItemDistinguishedName);
        }

        private GroupPrincipal GetSelectedGroupPrincipal()
        {
            return GroupPrincipal.FindByIdentity(
                _principalContext, _selectedItemDistinguishedName);
        }

        private UserPrincipal GetSelectedUserPrincipal()
        {
            return UserPrincipal.FindByIdentity(
                _principalContext, _selectedItemDistinguishedName);
        }

        private bool QueryTypeIsContextComputer()
        {
            return QueryType == ContextComputerGroups ||
                   QueryType == ContextComputerSummary;
        }

        private bool QueryTypeIsContextDirectReportOrUser()
        {
            return QueryType == ContextUserDirectReports ||
                   QueryType == ContextUserGroups ||
                   QueryType == ContextUserSummary ||
                   QueryType == ContextDirectReportDirectReports ||
                   QueryType == ContextDirectReportGroups ||
                   QueryType == ContextDirectReportSummary;
        }

        private bool QueryTypeIsContextGroup()
        {
            return QueryType == ContextGroupComputers ||
                   QueryType == ContextGroupUsers ||
                   QueryType == ContextGroupUsersDirectReports ||
                   QueryType == ContextGroupUsersGroups ||
                   QueryType == ContextGroupSummary;
        }

        private bool QueryTypeIsOu()
        {
            return QueryType == OuComputers ||
                   QueryType == OuGroups ||
                   QueryType == OuUsers ||
                   QueryType == OuUsersDirectReports ||
                   QueryType == OuUsersGroups;
        }

        private DataPreparer SetUpComputerDataPreparer()
        {
            var computerPrincipal = GetSelectedComputerPrincipal();
            Scope = computerPrincipal.Name;
            var computerDataPreparers = new Dictionary<QueryType, DataPreparer>
            {
                [ContextComputerGroups] = new DataPreparer
                {
                    Data = new[]
                    {
                        ActiveDirectorySearcher.GetComputerGroups(
                            computerPrincipal)
                    },
                    Attributes = DefaultComputerGroupsAttributes
                },
                [ContextComputerSummary] = new DataPreparer
                {
                    Data = new[]
                    {
                        computerPrincipal
                    },
                    Attributes = DefaultComputerAttributes
                }
            };
            return computerDataPreparers[QueryType];
        }

        private DataPreparer SetUpDirectReportOrUserDataPreparer()
        {
            var userPrincipal = GetSelectedUserPrincipal();
            Scope = userPrincipal.Name;
            var simplifiedQueryTypes =
                new Dictionary<QueryType, SimplifiedQueryType>
                {
                    [ContextDirectReportDirectReports] = DirectReports,
                    [ContextDirectReportGroups] = Groups,
                    [ContextDirectReportSummary] = Summary,
                    [ContextUserDirectReports] = DirectReports,
                    [ContextUserGroups] = Groups,
                    [ContextUserSummary] = Summary
                };
            var directReportOrUserDataPreparers = new Dictionary
                <SimplifiedQueryType, DataPreparer>
            {
                [DirectReports] = new DataPreparer
                {
                    Data = new[]
                    {
                        ActiveDirectorySearcher.GetUserDirectReports(
                            userPrincipal)
                    },
                    Attributes = DefaultUserDirectReportsAttributes
                },
                [Groups] = new DataPreparer
                {
                    Data = new[]
                    {
                        ActiveDirectorySearcher.GetUserGroups(userPrincipal)
                    },
                    Attributes = DefaultUserGroupsAttributes
                },
                [Summary] = new DataPreparer
                {
                    Data = new[]
                    {
                        userPrincipal
                    },
                    Attributes = DefaultUserAttributes
                }
            };
            return directReportOrUserDataPreparers[
                simplifiedQueryTypes[QueryType]];
        }

        private DataPreparer SetUpGroupDataPreparer()
        {
            var groupPrincipal = GetSelectedGroupPrincipal();
            Scope = groupPrincipal.Name;
            var groupDataPreparers = new Dictionary<QueryType, DataPreparer>
            {
                [ContextGroupComputers] = new DataPreparer
                {
                    Data = new[]
                    {
                        ActiveDirectorySearcher.GetComputerPrincipals(
                            groupPrincipal)
                    },
                    Attributes = DefaultGroupComputersAttributes
                },
                [ContextGroupSummary] = new DataPreparer
                {
                    Data = new[]
                    {
                        groupPrincipal
                    },
                    Attributes = DefaultGroupAttributes
                },
                [ContextGroupUsers] = new DataPreparer
                {
                    Data = new[]
                    {
                        ActiveDirectorySearcher.GetUserPrincipals(
                            groupPrincipal)
                    },
                    Attributes = DefaultGroupUsersAttributes
                },
                [ContextGroupUsersDirectReports] = new DataPreparer
                {
                    Data = new[]
                    {
                        ActiveDirectorySearcher
                            .GetGroupUsersDirectReports(
                                groupPrincipal, CancellationToken)
                    },
                    Attributes =
                        DefaultGroupUsersDirectReportsAttributes
                },
                [ContextGroupUsersGroups] = new DataPreparer
                {
                    Data = new[]
                    {
                        ActiveDirectorySearcher.GetGroupUsersGroups(
                            groupPrincipal, CancellationToken)
                    },
                    Attributes = DefaultGroupUsersGroupsAttributes
                }
            };
            return groupDataPreparers[QueryType];
        }

        private DataPreparer SetUpOuDataPreparer()
        {
            Scope = _activeDirectoryScope.Context;
            var activeDirectorySearcher = new ActiveDirectorySearcher(
                _activeDirectoryScope);
            var ouDataPreparers = new Dictionary<QueryType, DataPreparer>
            {
                [OuComputers] = new DataPreparer
                {
                    Data = activeDirectorySearcher.GetOuComputerPrincipals(),
                    Attributes = DefaultComputerAttributes
                },
                [OuGroups] = new DataPreparer
                {
                    Data = activeDirectorySearcher.GetOuGroupPrincipals(
                        CancellationToken),
                    Attributes = DefaultGroupAttributes
                },
                [OuUsers] = new DataPreparer
                {
                    Data = activeDirectorySearcher.GetOuUserPrincipals(),
                    Attributes = DefaultUserAttributes
                },
                [OuUsersDirectReports] = new DataPreparer
                {
                    Data = activeDirectorySearcher.GetOuUsersDirectReports(
                        CancellationToken),
                    Attributes = DefaultUserDirectReportsAttributes
                },
                [OuUsersGroups] = new DataPreparer
                {
                    Data = activeDirectorySearcher.GetOuUsersGroups(
                        CancellationToken),
                    Attributes = DefaultUserGroupsAttributes
                }
            };
            return ouDataPreparers[QueryType];
        }
    }
}