using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Threading;
using Extensions.PrincipalExtensions;
using static ActiveDirectoryTool.ActiveDirectoryProperty;

namespace ActiveDirectoryTool
{
    public class Searcher
    {
        private const char Asterix = '*';

        private static readonly ActiveDirectoryProperty[]
            DefaultComputerGroupsProperties =
            {
                ComputerName,
                GroupName,
                ComputerDistinguishedName,
                GroupDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultComputerProperties =
            {
                ComputerName,
                ComputerDescription,
                ComputerDistinguishedName
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
            DefaultGroupProperties =
            {
                GroupName,
                GroupManagedBy,
                GroupDescription,
                GroupDistinguishedName
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
            DefaultGroupUsersProperties =
            {
                ContainerGroupName,
                UserName,
                ContainerGroupDistinguishedName,
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

        private readonly CancellationToken _cancellationToken;
        private readonly IEnumerable<string> _distinguishedNames;
        private readonly QueryType _queryType;
        private readonly Scope _scope;
        private readonly string _searchText;
        private readonly DataPreparer _dataPreparer;

        public Searcher(
            QueryType queryType,
            Scope scope,
            IEnumerable<string> distinguishedNames,
            CancellationToken cancellationToken,
            string searchText)
        {
            _queryType = queryType;
            _scope = scope ?? Scope.GetDefaultScope();
            _distinguishedNames = distinguishedNames;
            _cancellationToken = cancellationToken;
            _searchText = searchText;
            _dataPreparer = new DataPreparer(cancellationToken);
        }

        public IEnumerable<ExpandoObject> GetData()
        {
            // ReSharper disable AccessToDisposedClosure
            IEnumerable<ExpandoObject> data;
            using (var principalContext = new PrincipalContext(
                ContextType.Domain,
                _scope.Domain,
                _scope.Context))
            {
                var mapping = new Dictionary
                    <QueryType, Func<IEnumerable<ExpandoObject>>>
                {
                    [QueryType.ComputersGroups] =
                        () => GetComputersGroupsData(),
                    [QueryType.ComputersSummaries] =
                        () => GetComputersSummariesData(),
                    [QueryType.DirectReportsDirectReports] =
                        () => GetUsersDirectReportsData(),
                    [QueryType.DirectReportsGroups] =
                        () => GetUsersGroupsData(),
                    [QueryType.DirectReportsSummaries] =
                        () => GetUsersSummariesData(),
                    [QueryType.GroupsComputers] =
                        () => GetGroupsComputersData(),
                    [QueryType.GroupsSummaries] =
                        () => GetGroupsSummariesData(),
                    [QueryType.GroupsUsers] = () => GetGroupsUsersData(),
                    [QueryType.GroupsUsersDirectReports] =
                        () => GetGroupsUsersDirectReportsData(),
                    [QueryType.GroupsUsersGroups] =
                        () => GetGroupsUsersGroupsData(),
                    [QueryType.OuComputers] = () => GetOuComputersData(
                        principalContext),
                    [QueryType.OuGroups] = () => GetOuGroupsData(
                        principalContext),
                    [QueryType.OuGroupsUsers] = () => GetOuGroupsUsersData(
                        principalContext),
                    [QueryType.OuUsers] = () => GetOuUsersData(
                        principalContext),
                    [QueryType.OuUsersDirectReports] =
                        () => GetOuUsersDirectReportsData(principalContext),
                    [QueryType.OuUsersGroups] = () => GetOuUsersGroupsData(
                        principalContext),
                    [QueryType.SearchComputer] = () => GetSearchComputerData(
                        principalContext),
                    [QueryType.SearchGroup] = () => GetSearchGroupData(
                        principalContext),
                    [QueryType.SearchUser] = () => GetSearchUserData(
                        principalContext),
                    [QueryType.UsersDirectReports] =
                        () => GetUsersDirectReportsData(),
                    [QueryType.UsersGroups] = () => GetUsersGroupsData(),
                    [QueryType.UsersSummaries] = () => GetUsersSummariesData()
                };
                data = mapping[_queryType]();
            }
            return data;
            // ReSharper restore AccessToDisposedClosure
        }

        private static PrincipalContext GetPrincipalContext()
        {
            return new PrincipalContext(ContextType.Domain);
        }

        private IEnumerable<ExpandoObject> GetComputersGroupsData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                using (var principalContext = GetPrincipalContext())
                using (var computerPrincipal = ComputerPrincipal
                    .FindByIdentity(
                        principalContext,
                        IdentityType.DistinguishedName,
                        distinguishedName))
                {
                    if (computerPrincipal == null) continue;
                    using (var groups = computerPrincipal.GetGroups())
                    {
                        foreach (var groupPrincipal in 
                            groups.GetGroupPrincipals())
                        {
                            _cancellationToken.ThrowIfCancellationRequested();
                            data.Add(
                                _dataPreparer.PrepareData(
                                    DefaultComputerGroupsProperties,
                                    computerPrincipal,
                                    groupPrincipal));
                        }
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetComputersSummariesData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                using (var principalContext = GetPrincipalContext())
                using (var computerPrincipal = ComputerPrincipal
                    .FindByIdentity(
                        principalContext,
                        IdentityType.DistinguishedName,
                        distinguishedName))
                {
                    if (computerPrincipal == null) continue;
                    data.Add(
                        _dataPreparer.PrepareData(
                            DefaultComputerProperties,
                            computerPrincipal));
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetGroupsComputersData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                using (var principalContext = GetPrincipalContext())
                using (var groupPrincipal = GroupPrincipal.FindByIdentity(
                    principalContext,
                    IdentityType.DistinguishedName,
                    distinguishedName))
                {
                    if (groupPrincipal == null) continue;
                    using (var members = groupPrincipal.GetMembers())
                    {
                        foreach (var computerPrincipal in members
                            .GetComputerPrincipals())
                        {
                            _cancellationToken.ThrowIfCancellationRequested();
                            data.Add(
                                _dataPreparer.PrepareData(
                                    DefaultGroupComputersProperties,
                                    containerGroupPrincipal: groupPrincipal,
                                    computerPrincipal: computerPrincipal));
                        }
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetGroupsSummariesData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                using (var principalContext = GetPrincipalContext())
                using (var groupPrincipal = GroupPrincipal.FindByIdentity(
                    principalContext,
                    IdentityType.DistinguishedName,
                    distinguishedName))
                {
                    if (groupPrincipal == null) continue;
                    data.Add(
                        _dataPreparer.PrepareData(
                            DefaultGroupProperties,
                            groupPrincipal: groupPrincipal));
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetGroupsUsersData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                using (var principalContext = GetPrincipalContext())
                using (var groupPrincipal = GroupPrincipal.FindByIdentity(
                    principalContext,
                    IdentityType.DistinguishedName,
                    distinguishedName))
                {
                    if (groupPrincipal == null) continue;
                    using (var members = groupPrincipal.GetMembers())
                    {
                        foreach (var userPrincipal in members
                            .GetUserPrincipals())
                        {
                            _cancellationToken.ThrowIfCancellationRequested();
                            data.Add(
                                _dataPreparer.PrepareData(
                                    DefaultGroupUsersDirectReportsProperties,
                                    containerGroupPrincipal: groupPrincipal,
                                    userPrincipal: userPrincipal));
                        }
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetGroupsUsersDirectReportsData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                using (var principalContext = GetPrincipalContext())
                using (var groupPrincipal = GroupPrincipal.FindByIdentity(
                    principalContext,
                    IdentityType.DistinguishedName,
                    distinguishedName))
                {
                    if (groupPrincipal == null) continue;
                    using (var members = groupPrincipal.GetMembers())
                    {
                        foreach (var userPrincipal in
                            members.GetUserPrincipals())
                        {
                            _cancellationToken.ThrowIfCancellationRequested();
                            if(userPrincipal == null) continue;
                            foreach (var directReportDistinguishedName in
                                userPrincipal
                                    .GetDirectReportDistinguishedNames())
                            {
                                _cancellationToken
                                    .ThrowIfCancellationRequested();
                                using (var directReportUserPrincipal =
                                    UserPrincipal.FindByIdentity(
                                        principalContext,
                                        IdentityType.DistinguishedName,
                                        directReportDistinguishedName))
                                {
                                    data.Add(
                                        _dataPreparer.PrepareData(
                                            DefaultGroupUsersDirectReportsProperties,
                                            containerGroupPrincipal:
                                                groupPrincipal,
                                            userPrincipal: userPrincipal,
                                            directReportUserPrincipal:
                                                directReportUserPrincipal));
                                }
                            }
                        }
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetGroupsUsersGroupsData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                using (var principalContext = GetPrincipalContext())
                using (var containerGroupPrincipal = GroupPrincipal
                    .FindByIdentity(
                        principalContext,
                        IdentityType.DistinguishedName,
                        distinguishedName))
                {
                    if (containerGroupPrincipal == null) continue;
                    using (var members = containerGroupPrincipal.GetMembers())
                    {
                        foreach (var userPrincipal in 
                            members.GetUserPrincipals())
                        {
                            _cancellationToken.ThrowIfCancellationRequested();
                            using (var groups = userPrincipal.GetGroups())
                            {
                                foreach (var groupPrincipal in groups
                                    .GetGroupPrincipals())
                                {
                                    _cancellationToken.ThrowIfCancellationRequested();
                                    data.Add(
                                        _dataPreparer.PrepareData(
                                            DefaultGroupUsersGroupsProperties,
                                            containerGroupPrincipal:
                                                containerGroupPrincipal,
                                            userPrincipal: userPrincipal,
                                            groupPrincipal: groupPrincipal));
                                }
                                
                            }
                        }
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetOuComputersData(
            PrincipalContext principalContext)
        {
            var data = new List<ExpandoObject>();
            using (var principal = new ComputerPrincipal(principalContext))
            using (var principalSearcher = new PrincipalSearcher(principal))
            using (var principalSearchResult = principalSearcher.FindAll())
            {
                foreach (var computerPrincipal in principalSearchResult
                    .GetComputerPrincipals())
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    data.Add(
                        _dataPreparer.PrepareData(
                            DefaultComputerProperties,
                            computerPrincipal));
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetOuGroupsData(
            PrincipalContext principalContext)
        {
            var data = new List<ExpandoObject>();
            using (var principal = new GroupPrincipal(principalContext))
            using (var principalSearcher = new PrincipalSearcher(principal))
            using (var principalSearchResult = principalSearcher.FindAll())
            {
                foreach (var groupPrincipal in principalSearchResult
                    .GetGroupPrincipals())
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    data.Add(
                        _dataPreparer.PrepareData(
                            DefaultGroupProperties,
                            groupPrincipal: groupPrincipal));
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetOuGroupsUsersData(
            PrincipalContext principalContext)
        {
            var data = new List<ExpandoObject>();
            using (var principal = new GroupPrincipal(principalContext))
            using (var principalSearcher = new PrincipalSearcher(principal))
            using (var principalSearchResult = principalSearcher.FindAll())
            {
                foreach (var groupPrincipal in principalSearchResult
                    .GetGroupPrincipals())
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    using (var members = groupPrincipal.GetMembers())
                    {
                        foreach (var userPrincipal in members
                            .GetUserPrincipals())
                        {
                            _cancellationToken.ThrowIfCancellationRequested();
                            data.Add(
                                _dataPreparer.PrepareData(
                                    DefaultGroupUsersProperties,
                                    containerGroupPrincipal: groupPrincipal,
                                    userPrincipal: userPrincipal));
                        }
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetOuUsersData(
            PrincipalContext principalContext)
        {
            var data = new List<ExpandoObject>();
            using (var principal = new UserPrincipal(principalContext))
            using (var principalSearcher = new PrincipalSearcher(principal))
            using (var principalSearchResult = principalSearcher.FindAll())
            {
                foreach (var userPrincipal in principalSearchResult
                    .GetUserPrincipals())
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    data.Add(
                        _dataPreparer.PrepareData(
                            DefaultUserProperties,
                            userPrincipal: userPrincipal));
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetOuUsersDirectReportsData(
            PrincipalContext principalContext)
        {
            var data = new List<ExpandoObject>();
            using (var principal = new UserPrincipal(principalContext))
            using (var principalSearcher = new PrincipalSearcher(principal))
            using (var principalSearchResult = principalSearcher.FindAll())
            {
                foreach (var userPrincipal in principalSearchResult
                    .GetUserPrincipals())
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    var directReportDistinguishedNames = userPrincipal
                        .GetDirectReportDistinguishedNames();
                    foreach (var directReportDistinguishedName in
                        directReportDistinguishedNames)
                    {
                        _cancellationToken.ThrowIfCancellationRequested();
                        using (var defaultPrincipalContext =
                            GetPrincipalContext())
                        using (var directReportUserPrincipal =
                            UserPrincipal.FindByIdentity(
                                defaultPrincipalContext,
                                IdentityType.DistinguishedName,
                                directReportDistinguishedName))
                        {
                            if (directReportUserPrincipal == null) continue;
                            data.Add(
                                _dataPreparer.PrepareData(
                                    DefaultUserDirectReportsProperties,
                                    userPrincipal: userPrincipal,
                                    directReportUserPrincipal:
                                        directReportUserPrincipal));
                        }
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetOuUsersGroupsData(
            PrincipalContext principalContext)
        {
            var data = new List<ExpandoObject>();
            using (var principal = new UserPrincipal(principalContext))
            using (var principalSearcher = new PrincipalSearcher(principal))
            using (var principalSearchResult = principalSearcher.FindAll())
            {
                foreach (var userPrincipal in principalSearchResult
                    .GetUserPrincipals())
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    using (var groups = userPrincipal.GetGroups())
                    {
                        foreach (var groupPrincipal in groups
                            .GetGroupPrincipals())
                        {
                            _cancellationToken.ThrowIfCancellationRequested();
                            data.Add(
                                _dataPreparer.PrepareData(
                                    DefaultUserGroupsProperties,
                                    userPrincipal: userPrincipal,
                                    groupPrincipal: groupPrincipal));
                        }
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetSearchComputerData(
            PrincipalContext principalContext)
        {
            var data = new List<ExpandoObject>();
            using (var principal = new ComputerPrincipal(principalContext)
            {
                Name = Asterix + _searchText + Asterix
            })
            using (var principalSearcher = new PrincipalSearcher(principal))
            using (var principalSearchResult = principalSearcher.FindAll())
            {
                foreach (var computerPrincipal in principalSearchResult
                    .GetComputerPrincipals())
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    data.Add(
                        _dataPreparer.PrepareData(
                            DefaultComputerProperties, computerPrincipal));
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetSearchUserData(
            PrincipalContext principalContext)
        {
            var data = new List<ExpandoObject>();
            using (var principal = new UserPrincipal(principalContext)
            {
                Name = Asterix + _searchText + Asterix
            })
            using (var principalSearcher = new PrincipalSearcher(principal))
            using (var principalSearchResult = principalSearcher.FindAll())
            {
                foreach (var userPrincipal in principalSearchResult
                    .GetUserPrincipals())
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    data.Add(
                        _dataPreparer.PrepareData(
                            DefaultUserProperties, 
                            userPrincipal: userPrincipal));
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetSearchGroupData(
            PrincipalContext principalContext)
        {
            var data = new List<ExpandoObject>();
            using (var principal = new GroupPrincipal(principalContext)
            {
                Name = Asterix + _searchText + Asterix
            })
            using (var principalSearcher = new PrincipalSearcher(principal))
            using (var principalSearchResult = principalSearcher.FindAll())
            {
                foreach (var groupPrincipal in principalSearchResult
                    .GetGroupPrincipals())
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    data.Add(
                        _dataPreparer.PrepareData(
                            DefaultGroupProperties,
                            groupPrincipal: groupPrincipal));
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetUsersDirectReportsData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                using (var principalContext = GetPrincipalContext())
                using (var userPrincipal = UserPrincipal.FindByIdentity(
                    principalContext,
                    IdentityType.DistinguishedName,
                    distinguishedName))
                {
                    if (userPrincipal == null) continue;
                    var directReportDistinguishedNames =
                        userPrincipal.GetDirectReportDistinguishedNames();
                    foreach (var directReportDistinguishedName in 
                        directReportDistinguishedNames)
                    {
                        _cancellationToken.ThrowIfCancellationRequested();
                        using (var directReportUserPrincipal = UserPrincipal
                            .FindByIdentity(
                                principalContext,
                                IdentityType.DistinguishedName,
                                directReportDistinguishedName))
                        {
                            data.Add(
                                _dataPreparer.PrepareData(
                                    DefaultUserDirectReportsProperties,
                                    userPrincipal: userPrincipal,
                                    directReportUserPrincipal:
                                        directReportUserPrincipal));
                        }
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetUsersGroupsData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                using (var principalContext = GetPrincipalContext())
                using (var userPrincipal = UserPrincipal.FindByIdentity(
                    principalContext,
                    IdentityType.DistinguishedName,
                    distinguishedName))
                {
                    if (userPrincipal == null) continue;
                    using (var groups = userPrincipal.GetGroups())
                    {
                        foreach (var groupPrincipal in groups
                            .GetGroupPrincipals())
                        {
                            _cancellationToken.ThrowIfCancellationRequested();
                            data.Add(
                                _dataPreparer.PrepareData(
                                    DefaultUserGroupsProperties,
                                    userPrincipal: userPrincipal,
                                    groupPrincipal: groupPrincipal));
                        }
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetUsersSummariesData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                using (var principalContext = GetPrincipalContext())
                using (var userPrincipal = UserPrincipal.FindByIdentity(
                    principalContext,
                    IdentityType.DistinguishedName,
                    distinguishedName))
                {
                    if (userPrincipal == null) continue;
                    data.Add(
                        _dataPreparer.PrepareData(
                            DefaultUserProperties,
                            userPrincipal: userPrincipal));
                }
            }
            return data;
        }
    }
}