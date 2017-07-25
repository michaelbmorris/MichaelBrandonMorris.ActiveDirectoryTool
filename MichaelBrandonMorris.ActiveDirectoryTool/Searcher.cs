using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Threading;
using MichaelBrandonMorris.Extensions.PrincipalExtensions;

namespace MichaelBrandonMorris.ActiveDirectoryTool
{
    /// <summary>
    ///     Class Searcher.
    /// </summary>
    /// TODO Edit XML Comment Template for Searcher
    public class Searcher
    {
        /// <summary>
        ///     The asterisk
        /// </summary>
        /// TODO Edit XML Comment Template for Asterisk
        private const char Asterisk = '*';

        /// <summary>
        ///     The default computer groups properties
        /// </summary>
        /// TODO Edit XML Comment Template for DefaultComputerGroupsProperties
        private static readonly ActiveDirectoryProperty[]
            DefaultComputerGroupsProperties =
            {
                ActiveDirectoryProperty.ComputerName,
                ActiveDirectoryProperty.GroupName,
                ActiveDirectoryProperty.ComputerDistinguishedName,
                ActiveDirectoryProperty.GroupDistinguishedName
            };

        /// <summary>
        ///     The default computer properties
        /// </summary>
        /// TODO Edit XML Comment Template for DefaultComputerProperties
        private static readonly ActiveDirectoryProperty[]
            DefaultComputerProperties =
            {
                ActiveDirectoryProperty.ComputerName,
                ActiveDirectoryProperty.ComputerDescription,
                ActiveDirectoryProperty.ComputerDistinguishedName
            };

        /// <summary>
        ///     The default group computers properties
        /// </summary>
        /// TODO Edit XML Comment Template for DefaultGroupComputersProperties
        private static readonly ActiveDirectoryProperty[]
            DefaultGroupComputersProperties =
            {
                ActiveDirectoryProperty.GroupName,
                ActiveDirectoryProperty.ComputerName,
                ActiveDirectoryProperty.GroupDistinguishedName,
                ActiveDirectoryProperty.ComputerDistinguishedName
            };

        /// <summary>
        ///     The default group properties
        /// </summary>
        /// TODO Edit XML Comment Template for DefaultGroupProperties
        private static readonly ActiveDirectoryProperty[] DefaultGroupProperties
            =
            {
                ActiveDirectoryProperty.GroupName,
                ActiveDirectoryProperty.GroupManagedByName,
                ActiveDirectoryProperty.GroupDescription,
                ActiveDirectoryProperty.GroupDistinguishedName,
                ActiveDirectoryProperty.GroupManagedByDistinguishedName
            };

        /// <summary>
        ///     The default group users direct reports properties
        /// </summary>
        /// TODO Edit XML Comment Template for DefaultGroupUsersDirectReportsProperties
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

        /// <summary>
        ///     The default group users groups properties
        /// </summary>
        /// TODO Edit XML Comment Template for DefaultGroupUsersGroupsProperties
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

        /// <summary>
        ///     The default group users properties
        /// </summary>
        /// TODO Edit XML Comment Template for DefaultGroupUsersProperties
        private static readonly ActiveDirectoryProperty[]
            DefaultGroupUsersProperties =
            {
                ActiveDirectoryProperty.ContainerGroupName,
                ActiveDirectoryProperty.UserName,
                ActiveDirectoryProperty.ContainerGroupDistinguishedName,
                ActiveDirectoryProperty.UserDistinguishedName
            };

        /// <summary>
        ///     The default user direct reports properties
        /// </summary>
        /// TODO Edit XML Comment Template for DefaultUserDirectReportsProperties
        private static readonly ActiveDirectoryProperty[]
            DefaultUserDirectReportsProperties =
            {
                ActiveDirectoryProperty.UserName,
                ActiveDirectoryProperty.DirectReportName,
                ActiveDirectoryProperty.UserDistinguishedName,
                ActiveDirectoryProperty.DirectReportDistinguishedName
            };

        /// <summary>
        ///     The default user groups properties
        /// </summary>
        /// TODO Edit XML Comment Template for DefaultUserGroupsProperties
        private static readonly ActiveDirectoryProperty[]
            DefaultUserGroupsProperties =
            {
                ActiveDirectoryProperty.UserName,
                ActiveDirectoryProperty.GroupName,
                ActiveDirectoryProperty.UserDistinguishedName,
                ActiveDirectoryProperty.GroupDistinguishedName
            };

        /// <summary>
        ///     The default user properties
        /// </summary>
        /// TODO Edit XML Comment Template for DefaultUserProperties
        private static readonly ActiveDirectoryProperty[] DefaultUserProperties
            =
            {
                ActiveDirectoryProperty.UserSurname,
                ActiveDirectoryProperty.UserGivenName,
                ActiveDirectoryProperty.UserDisplayName,
                ActiveDirectoryProperty.UserSamAccountName,
                ActiveDirectoryProperty.UserIsActive,
                ActiveDirectoryProperty.UserAccountExpirationDate,
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
                ActiveDirectoryProperty.UserEmployeeNumber,
                ActiveDirectoryProperty.UserEmployeeNumberHash,
                ActiveDirectoryProperty.UserDistinguishedName,
                ActiveDirectoryProperty.ManagerDistinguishedName
            };

        /// <summary>
        ///     Initializes a new instance of the
        ///     <see cref="Searcher" /> class.
        /// </summary>
        /// <param name="queryType">Type of the query.</param>
        /// <param name="scope">The scope.</param>
        /// <param name="distinguishedNames">The distinguished names.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <param name="searchText">The search text.</param>
        /// TODO Edit XML Comment Template for #ctor
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

        /// <summary>
        ///     Gets the cancellation token.
        /// </summary>
        /// <value>The cancellation token.</value>
        /// TODO Edit XML Comment Template for CancellationToken
        private CancellationToken CancellationToken
        {
            get;
        }

        /// <summary>
        ///     Gets the data preparer.
        /// </summary>
        /// <value>The data preparer.</value>
        /// TODO Edit XML Comment Template for DataPreparer
        private DataPreparer DataPreparer
        {
            get;
        }

        /// <summary>
        ///     Gets the distinguished names.
        /// </summary>
        /// <value>The distinguished names.</value>
        /// TODO Edit XML Comment Template for DistinguishedNames
        private IEnumerable<string> DistinguishedNames
        {
            get;
        }

        /// <summary>
        ///     Gets the type of the query.
        /// </summary>
        /// <value>The type of the query.</value>
        /// TODO Edit XML Comment Template for QueryType
        private QueryType QueryType
        {
            get;
        }

        /// <summary>
        ///     Gets the scope.
        /// </summary>
        /// <value>The scope.</value>
        /// TODO Edit XML Comment Template for Scope
        private Scope Scope
        {
            get;
        }

        /// <summary>
        ///     Gets the search text.
        /// </summary>
        /// <value>The search text.</value>
        /// TODO Edit XML Comment Template for SearchText
        private string SearchText
        {
            get;
        }

        /// <summary>
        ///     Gets the data.
        /// </summary>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetData
        public IEnumerable<ExpandoObject> GetData()
        {
            // ReSharper disable AccessToDisposedClosure
            IEnumerable<ExpandoObject> data;
            using (var principalContext =
                Scope == null
                    ? GetPrincipalContext()
                    : new PrincipalContext(
                        ContextType.Domain,
                        Scope.Domain,
                        Scope.Context))
            {
                var mapping =
                    new Dictionary<QueryType, Func<IEnumerable<ExpandoObject>>
                    >
                    {
                        [QueryType.ComputersGroups] =
                        () => GetComputersGroupsData(),
                        [QueryType.ComputersSummaries] =
                        () => GetComputersSummariesData(),
                        [QueryType.GroupsComputers] =
                        () => GetGroupsComputersData(),
                        [QueryType.GroupsSummaries] =
                        () => GetGroupsSummariesData(),
                        [QueryType.GroupsUsers] = () => GetGroupsUsersData(),
                        [QueryType.GroupsUsersDirectReports] =
                        () => GetGroupsUsersDirectReportsData(),
                        [QueryType.GroupsUsersGroups] =
                        () => GetGroupsUsersGroupsData(),
                        [QueryType.OuComputers] =
                        () => GetOuComputersData(principalContext),
                        [QueryType.OuGroups] =
                        () => GetOuGroupsData(principalContext),
                        [QueryType.OuGroupsUsers] =
                        () => GetOuGroupsUsersData(principalContext),
                        [QueryType.OuUsers] =
                        () => GetOuUsersData(principalContext),
                        [QueryType.OuUsersDirectReports] =
                        () => GetOuUsersDirectReportsData(principalContext),
                        [QueryType.OuUsersGroups] =
                        () => GetOuUsersGroupsData(principalContext),
                        [QueryType.SearchComputer] =
                        () => GetSearchComputerData(principalContext),
                        [QueryType.SearchGroup] =
                        () => GetSearchGroupData(principalContext),
                        [QueryType.SearchUser] =
                        () => GetSearchUserData(principalContext),
                        [QueryType.UsersDirectReports] =
                        () => GetUsersDirectReportsData(),
                        [QueryType.UsersGroups] = () => GetUsersGroupsData(),
                        [QueryType.UsersSummaries] =
                        () => GetUsersSummariesData()
                    };
                data = mapping[QueryType]();
            }
            return data;
            // ReSharper restore AccessToDisposedClosure
        }

        /// <summary>
        ///     Gets the principal context.
        /// </summary>
        /// <returns>PrincipalContext.</returns>
        /// TODO Edit XML Comment Template for GetPrincipalContext
        private static PrincipalContext GetPrincipalContext()
        {
            return new PrincipalContext(ContextType.Domain);
        }

        /// <summary>
        ///     Gets the computers groups data.
        /// </summary>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetComputersGroupsData
        private IEnumerable<ExpandoObject> GetComputersGroupsData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in DistinguishedNames)
            {
                CancellationToken.ThrowIfCancellationRequested();
                using (var principalContext = GetPrincipalContext())
                using (var computerPrincipal = ComputerPrincipal.FindByIdentity(
                    principalContext,
                    IdentityType.DistinguishedName,
                    distinguishedName))
                {
                    if (computerPrincipal == null)
                    {
                        continue;
                    }

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

        /// <summary>
        ///     Gets the computers summaries data.
        /// </summary>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetComputersSummariesData
        private IEnumerable<ExpandoObject> GetComputersSummariesData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in DistinguishedNames)
            {
                CancellationToken.ThrowIfCancellationRequested();
                using (var principalContext = GetPrincipalContext())
                using (var computerPrincipal = ComputerPrincipal.FindByIdentity(
                    principalContext,
                    IdentityType.DistinguishedName,
                    distinguishedName))
                {
                    if (computerPrincipal == null)
                    {
                        continue;
                    }

                    data.Add(
                        DataPreparer.PrepareData(
                            DefaultComputerProperties,
                            computerPrincipal));
                }
            }

            return data;
        }

        /// <summary>
        ///     Gets the groups computers data.
        /// </summary>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetGroupsComputersData
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
                    if (groupPrincipal == null)
                    {
                        continue;
                    }

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

        /// <summary>
        ///     Gets the groups summaries data.
        /// </summary>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetGroupsSummariesData
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
                    if (groupPrincipal == null)
                    {
                        continue;
                    }

                    data.Add(
                        DataPreparer.PrepareData(
                            DefaultGroupProperties,
                            groupPrincipal: groupPrincipal));
                }
            }

            return data;
        }

        /// <summary>
        ///     Gets the groups users data.
        /// </summary>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetGroupsUsersData
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
                    if (groupPrincipal == null)
                    {
                        continue;
                    }

                    using (var members = groupPrincipal.GetMembers())
                    {
                        foreach (var userPrincipal in
                            members.GetUserPrincipals())
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

        /// <summary>
        ///     Gets the groups users direct reports data.
        /// </summary>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetGroupsUsersDirectReportsData
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
                    if (groupPrincipal == null)
                    {
                        continue;
                    }

                    using (var members = groupPrincipal.GetMembers())
                    {
                        foreach (var userPrincipal in
                            members.GetUserPrincipals())
                        {
                            CancellationToken.ThrowIfCancellationRequested();
                            if (userPrincipal == null)
                            {
                                continue;
                            }

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

        /// <summary>
        ///     Gets the groups users groups data.
        /// </summary>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetGroupsUsersGroupsData
        private IEnumerable<ExpandoObject> GetGroupsUsersGroupsData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in DistinguishedNames)
            {
                CancellationToken.ThrowIfCancellationRequested();
                using (var principalContext = GetPrincipalContext())
                using (var containerGroupPrincipal =
                    GroupPrincipal.FindByIdentity(
                        principalContext,
                        IdentityType.DistinguishedName,
                        distinguishedName))
                {
                    if (containerGroupPrincipal == null)
                    {
                        continue;
                    }

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

        /// <summary>
        ///     Gets the ou computers data.
        /// </summary>
        /// <param name="principalContext">The principal context.</param>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetOuComputersData
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

        /// <summary>
        ///     Gets the ou groups data.
        /// </summary>
        /// <param name="principalContext">The principal context.</param>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetOuGroupsData
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

        /// <summary>
        ///     Gets the ou groups users data.
        /// </summary>
        /// <param name="principalContext">The principal context.</param>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetOuGroupsUsersData
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
                        foreach (var userPrincipal in
                            members.GetUserPrincipals())
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

        /// <summary>
        ///     Gets the ou users data.
        /// </summary>
        /// <param name="principalContext">The principal context.</param>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetOuUsersData
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

        /// <summary>
        ///     Gets the ou users direct reports data.
        /// </summary>
        /// <param name="principalContext">The principal context.</param>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetOuUsersDirectReportsData
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
                            if (directReportUserPrincipal == null)
                            {
                                continue;
                            }

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

        /// <summary>
        ///     Gets the ou users groups data.
        /// </summary>
        /// <param name="principalContext">The principal context.</param>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetOuUsersGroupsData
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
                        foreach (var groupPrincipal in
                            groups.GetGroupPrincipals())
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

        /// <summary>
        ///     Gets the search computer data.
        /// </summary>
        /// <param name="principalContext">The principal context.</param>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetSearchComputerData
        private IEnumerable<ExpandoObject> GetSearchComputerData(
            PrincipalContext principalContext)
        {
            var data = new List<ExpandoObject>();
            using (var principal = new ComputerPrincipal(principalContext))
            {
                principal.Name = Asterisk + SearchText + Asterisk;
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
            }

            return data;
        }

        /// <summary>
        ///     Gets the search group data.
        /// </summary>
        /// <param name="principalContext">The principal context.</param>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetSearchGroupData
        private IEnumerable<ExpandoObject> GetSearchGroupData(
            PrincipalContext principalContext)
        {
            var data = new List<ExpandoObject>();
            using (var principal = new GroupPrincipal(principalContext))
            {
                principal.Name = Asterisk + SearchText + Asterisk;
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
            }

            return data;
        }

        /// <summary>
        ///     Gets the search user data.
        /// </summary>
        /// <param name="principalContext">The principal context.</param>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetSearchUserData
        private IEnumerable<ExpandoObject> GetSearchUserData(
            PrincipalContext principalContext)
        {
            var data = new List<ExpandoObject>();
            using (var principal = new UserPrincipal(principalContext))
            {
                principal.Name = Asterisk + SearchText + Asterisk;
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
            }

            return data;
        }

        /// <summary>
        ///     Gets the users direct reports data.
        /// </summary>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetUsersDirectReportsData
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
                    if (userPrincipal == null)
                    {
                        continue;
                    }

                    var directReportDistinguishedNames = userPrincipal
                        .GetDirectReportDistinguishedNames();
                    foreach (var directReportDistinguishedName in
                        directReportDistinguishedNames)
                    {
                        CancellationToken.ThrowIfCancellationRequested();
                        using (var directReportUserPrincipal =
                            UserPrincipal.FindByIdentity(
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

        /// <summary>
        ///     Gets the users groups data.
        /// </summary>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetUsersGroupsData
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
                    if (userPrincipal == null)
                    {
                        continue;
                    }

                    using (var groups = userPrincipal.GetGroups())
                    {
                        foreach (var groupPrincipal in
                            groups.GetGroupPrincipals())
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

        /// <summary>
        ///     Gets the users summaries data.
        /// </summary>
        /// <returns>IEnumerable&lt;ExpandoObject&gt;.</returns>
        /// TODO Edit XML Comment Template for GetUsersSummariesData
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
                    if (userPrincipal == null)
                    {
                        continue;
                    }

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