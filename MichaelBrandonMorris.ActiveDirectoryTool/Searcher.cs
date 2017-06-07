using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Threading;
using MichaelBrandonMorris.Extensions.PrincipalExtensions;

namespace MichaelBrandonMorris.ActiveDirectoryTool
{
    public class Searcher
    {
        private const char Asterix = '*';

        private static readonly ActiveDirectoryProperty[]
            DefaultComputerGroupsProperties =
            {
                ActiveDirectoryProperty.ComputerName,
                ActiveDirectoryProperty.GroupName,
                ActiveDirectoryProperty.ComputerDistinguishedName,
                ActiveDirectoryProperty.GroupDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultComputerProperties =
            {
                ActiveDirectoryProperty.ComputerName,
                ActiveDirectoryProperty.ComputerDescription,
                ActiveDirectoryProperty.ComputerDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultGroupComputersProperties =
            {
                ActiveDirectoryProperty.GroupName,
                ActiveDirectoryProperty.ComputerName,
                ActiveDirectoryProperty.GroupDistinguishedName,
                ActiveDirectoryProperty.ComputerDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultGroupProperties =
            {
                ActiveDirectoryProperty.GroupName,
                ActiveDirectoryProperty.GroupManagedByName,
                ActiveDirectoryProperty.GroupDescription,
                ActiveDirectoryProperty.GroupDistinguishedName,
                ActiveDirectoryProperty.GroupManagedByDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultGroupUsersDirectReportsProperties =
            {
                ActiveDirectoryProperty.ContainerGroupName,
                ActiveDirectoryProperty.UserName,
                ActiveDirectoryProperty.DirectReportName,
                ActiveDirectoryProperty.ContainerGroupDistinguishedName,
                ActiveDirectoryProperty.UserDistinguishedName,
                ActiveDirectoryProperty.DirectReportDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultGroupUsersGroupsProperties =
            {
                ActiveDirectoryProperty.ContainerGroupName,
                ActiveDirectoryProperty.UserName,
                ActiveDirectoryProperty.GroupName,
                ActiveDirectoryProperty.ContainerGroupDistinguishedName,
                ActiveDirectoryProperty.UserDistinguishedName,
                ActiveDirectoryProperty.GroupDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultGroupUsersProperties =
            {
                ActiveDirectoryProperty.ContainerGroupName,
                ActiveDirectoryProperty.UserName,
                ActiveDirectoryProperty.ContainerGroupDistinguishedName,
                ActiveDirectoryProperty.UserDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultUserDirectReportsProperties =
            {
                ActiveDirectoryProperty.UserName,
                ActiveDirectoryProperty.DirectReportName,
                ActiveDirectoryProperty.UserDistinguishedName,
                ActiveDirectoryProperty.DirectReportDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultUserGroupsProperties =
            {
                ActiveDirectoryProperty.UserName,
                ActiveDirectoryProperty.GroupName,
                ActiveDirectoryProperty.UserDistinguishedName,
                ActiveDirectoryProperty.GroupDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultUserProperties =
            {
                ActiveDirectoryProperty.UserSurname,
                ActiveDirectoryProperty.UserGivenName,
                ActiveDirectoryProperty.UserDisplayName,
                ActiveDirectoryProperty.UserSamAccountName,
                ActiveDirectoryProperty.UserIsActive,
                ActiveDirectoryProperty.UserIsAccountLockedOut,
                ActiveDirectoryProperty.UserLastLogon,
                ActiveDirectoryProperty.UserDescription,
                ActiveDirectoryProperty.UserTitle,
                ActiveDirectoryProperty.UserCompany,
                ActiveDirectoryProperty.ManagerName,
                ActiveDirectoryProperty.UserHomeDrive,
                ActiveDirectoryProperty.UserHomeDirectory,
                ActiveDirectoryProperty.UserScriptPath,
                ActiveDirectoryProperty.UserEmailAddress,
                ActiveDirectoryProperty.UserStreetAddress,
                ActiveDirectoryProperty.UserCity,
                ActiveDirectoryProperty.UserState,
                ActiveDirectoryProperty.UserVoiceTelephoneNumber,
                ActiveDirectoryProperty.UserPager,
                ActiveDirectoryProperty.UserMobile,
                ActiveDirectoryProperty.UserFax,
                ActiveDirectoryProperty.UserVoip,
                ActiveDirectoryProperty.UserSip,
                ActiveDirectoryProperty.UserUserPrincipalName,
                ActiveDirectoryProperty.UserDistinguishedName,
                ActiveDirectoryProperty.ManagerDistinguishedName
            };

        public Searcher(
            QueryType queryType,
            Scope scope,
            IEnumerable<string> distinguishedNames,
            CancellationToken cancellationToken,
            string searchText)
        {
            QueryType = queryType;
            Scope = scope;
            DistinguishedNames = distinguishedNames;
            CancellationToken = cancellationToken;
            SearchText = searchText;
            DataPreparer = new DataPreparer(cancellationToken);
        }

        private CancellationToken CancellationToken
        {
            get;
        }

        private DataPreparer DataPreparer
        {
            get;
        }

        private IEnumerable<string> DistinguishedNames
        {
            get;
        }

        private QueryType QueryType
        {
            get;
        }

        private Scope Scope
        {
            get;
        }

        private string SearchText
        {
            get;
        }

        public IEnumerable<ExpandoObject> GetData()
        {
            // ReSharper disable AccessToDisposedClosure
            IEnumerable<ExpandoObject> data;
            using (var principalContext = Scope == null
                ? GetPrincipalContext()
                : new PrincipalContext(
                    ContextType.Domain,
                    Scope.Domain,
                    Scope.Context))
            {
                var mapping = new Dictionary
                    <QueryType, Func<IEnumerable<ExpandoObject>>>
                {
                    [QueryType.ComputersGroups] = () =>
                        GetComputersGroupsData(),
                    [QueryType.ComputersSummaries] =
                        () =>
                            GetComputersSummariesData(),
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
                    [QueryType.OuUsers] = () =>
                        GetOuUsersData(principalContext),
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
                data = mapping[QueryType]();
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
            foreach (var distinguishedName in DistinguishedNames)
            {
                CancellationToken.ThrowIfCancellationRequested();
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
                            CancellationToken.ThrowIfCancellationRequested();
                            data.Add(
                                DataPreparer.PrepareData(
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
            foreach (var distinguishedName in DistinguishedNames)
            {
                CancellationToken.ThrowIfCancellationRequested();
                using (var principalContext = GetPrincipalContext())
                using (var computerPrincipal = ComputerPrincipal
                    .FindByIdentity(
                        principalContext,
                        IdentityType.DistinguishedName,
                        distinguishedName))
                {
                    if (computerPrincipal == null) continue;
                    data.Add(
                        DataPreparer.PrepareData(
                            DefaultComputerProperties,
                            computerPrincipal));
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetGroupsComputersData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in DistinguishedNames)
            {
                CancellationToken.ThrowIfCancellationRequested();
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
                            CancellationToken.ThrowIfCancellationRequested();
                            data.Add(
                                DataPreparer.PrepareData(
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
            foreach (var distinguishedName in DistinguishedNames)
            {
                using (var principalContext = GetPrincipalContext())
                using (var groupPrincipal = GroupPrincipal.FindByIdentity(
                    principalContext,
                    IdentityType.DistinguishedName,
                    distinguishedName))
                {
                    if (groupPrincipal == null) continue;
                    data.Add(
                        DataPreparer.PrepareData(
                            DefaultGroupProperties,
                            groupPrincipal: groupPrincipal));
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetGroupsUsersData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in DistinguishedNames)
            {
                CancellationToken.ThrowIfCancellationRequested();
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
                            CancellationToken.ThrowIfCancellationRequested();
                            data.Add(
                                DataPreparer.PrepareData(
                                    DefaultGroupUsersProperties,
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
            foreach (var distinguishedName in DistinguishedNames)
            {
                CancellationToken.ThrowIfCancellationRequested();
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
                            CancellationToken.ThrowIfCancellationRequested();
                            if (userPrincipal == null) continue;
                            foreach (var directReportDistinguishedName in
                                userPrincipal
                                    .GetDirectReportDistinguishedNames())
                            {
                                CancellationToken
                                    .ThrowIfCancellationRequested();
                                using (var directReportUserPrincipal =
                                    UserPrincipal.FindByIdentity(
                                        principalContext,
                                        IdentityType.DistinguishedName,
                                        directReportDistinguishedName))
                                {
                                    data.Add(
                                        DataPreparer.PrepareData(
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
            foreach (var distinguishedName in DistinguishedNames)
            {
                CancellationToken.ThrowIfCancellationRequested();
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
                            CancellationToken.ThrowIfCancellationRequested();
                            using (var groups = userPrincipal.GetGroups())
                            {
                                foreach (var groupPrincipal in groups
                                    .GetGroupPrincipals())
                                {
                                    CancellationToken
                                        .ThrowIfCancellationRequested();
                                    data.Add(
                                        DataPreparer.PrepareData(
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
                    CancellationToken.ThrowIfCancellationRequested();
                    data.Add(
                        DataPreparer.PrepareData(
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
                    CancellationToken.ThrowIfCancellationRequested();
                    data.Add(
                        DataPreparer.PrepareData(
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
                    CancellationToken.ThrowIfCancellationRequested();
                    using (var members = groupPrincipal.GetMembers())
                    {
                        foreach (var userPrincipal in members
                            .GetUserPrincipals())
                        {
                            CancellationToken.ThrowIfCancellationRequested();
                            data.Add(
                                DataPreparer.PrepareData(
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
                    CancellationToken.ThrowIfCancellationRequested();
                    data.Add(
                        DataPreparer.PrepareData(
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
                    CancellationToken.ThrowIfCancellationRequested();
                    var directReportDistinguishedNames = userPrincipal
                        .GetDirectReportDistinguishedNames();
                    foreach (var directReportDistinguishedName in
                        directReportDistinguishedNames)
                    {
                        CancellationToken.ThrowIfCancellationRequested();
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
                                DataPreparer.PrepareData(
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
                    CancellationToken.ThrowIfCancellationRequested();
                    using (var groups = userPrincipal.GetGroups())
                    {
                        foreach (var groupPrincipal in groups
                            .GetGroupPrincipals())
                        {
                            CancellationToken.ThrowIfCancellationRequested();
                            data.Add(
                                DataPreparer.PrepareData(
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
            using (var principal = new ComputerPrincipal(principalContext))
            {
                principal.Name = Asterix + SearchText + Asterix;
                using (var principalSearcher = new PrincipalSearcher(
                    principal))
                using (var principalSearchResult = principalSearcher.FindAll())
                {
                    foreach (var computerPrincipal in principalSearchResult
                        .GetComputerPrincipals())
                    {
                        CancellationToken.ThrowIfCancellationRequested();
                        data.Add(
                            DataPreparer.PrepareData(
                                DefaultComputerProperties, computerPrincipal));
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetSearchGroupData(
            PrincipalContext principalContext)
        {
            var data = new List<ExpandoObject>();
            using (var principal = new GroupPrincipal(principalContext))
            {
                principal.Name = Asterix + SearchText + Asterix;
                using (var principalSearcher = new PrincipalSearcher(
                    principal))
                using (var principalSearchResult = principalSearcher.FindAll())
                {
                    foreach (var groupPrincipal in principalSearchResult
                        .GetGroupPrincipals())
                    {
                        CancellationToken.ThrowIfCancellationRequested();
                        data.Add(
                            DataPreparer.PrepareData(
                                DefaultGroupProperties,
                                groupPrincipal: groupPrincipal));
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetSearchUserData(
            PrincipalContext principalContext)
        {
            var data = new List<ExpandoObject>();
            using (var principal = new UserPrincipal(principalContext))
            {
                principal.Name = Asterix + SearchText + Asterix;
                using (var principalSearcher = new PrincipalSearcher(
                    principal))
                using (var principalSearchResult = principalSearcher.FindAll())
                {
                    foreach (var userPrincipal in principalSearchResult
                        .GetUserPrincipals())
                    {
                        CancellationToken.ThrowIfCancellationRequested();
                        data.Add(
                            DataPreparer.PrepareData(
                                DefaultUserProperties,
                                userPrincipal: userPrincipal));
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetUsersDirectReportsData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in DistinguishedNames)
            {
                CancellationToken.ThrowIfCancellationRequested();
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
                        CancellationToken.ThrowIfCancellationRequested();
                        using (var directReportUserPrincipal = UserPrincipal
                            .FindByIdentity(
                                principalContext,
                                IdentityType.DistinguishedName,
                                directReportDistinguishedName))
                        {
                            data.Add(
                                DataPreparer.PrepareData(
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
            foreach (var distinguishedName in DistinguishedNames)
            {
                CancellationToken.ThrowIfCancellationRequested();
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
                            CancellationToken.ThrowIfCancellationRequested();
                            data.Add(
                                DataPreparer.PrepareData(
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
            foreach (var distinguishedName in DistinguishedNames)
            {
                CancellationToken.ThrowIfCancellationRequested();
                using (var principalContext = GetPrincipalContext())
                using (var userPrincipal = UserPrincipal.FindByIdentity(
                    principalContext,
                    IdentityType.DistinguishedName,
                    distinguishedName))
                {
                    if (userPrincipal == null) continue;
                    data.Add(
                        DataPreparer.PrepareData(
                            DefaultUserProperties,
                            userPrincipal: userPrincipal));
                }
            }
            return data;
        }
    }
}