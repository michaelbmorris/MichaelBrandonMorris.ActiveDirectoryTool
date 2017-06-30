using System.Collections.Generic;
using System.Dynamic;
using System.Threading;
using System.Threading.Tasks;

namespace MichaelBrandonMorris.ActiveDirectoryTool
{
    /// <summary>
    ///     Enum QueryType
    /// </summary>
    /// TODO Edit XML Comment Template for QueryType
    public enum QueryType
    {
        /// <summary>
        ///     The none
        /// </summary>
        /// TODO Edit XML Comment Template for None
        None,

        /// <summary>
        ///     The computers groups
        /// </summary>
        /// TODO Edit XML Comment Template for ComputersGroups
        ComputersGroups,

        /// <summary>
        ///     The computers summaries
        /// </summary>
        /// TODO Edit XML Comment Template for ComputersSummaries
        ComputersSummaries,

        /// <summary>
        ///     The groups computers
        /// </summary>
        /// TODO Edit XML Comment Template for GroupsComputers
        GroupsComputers,

        /// <summary>
        ///     The groups users
        /// </summary>
        /// TODO Edit XML Comment Template for GroupsUsers
        GroupsUsers,

        /// <summary>
        ///     The groups users direct reports
        /// </summary>
        /// TODO Edit XML Comment Template for GroupsUsersDirectReports
        GroupsUsersDirectReports,

        /// <summary>
        ///     The groups users groups
        /// </summary>
        /// TODO Edit XML Comment Template for GroupsUsersGroups
        GroupsUsersGroups,

        /// <summary>
        ///     The groups summaries
        /// </summary>
        /// TODO Edit XML Comment Template for GroupsSummaries
        GroupsSummaries,

        /// <summary>
        ///     The users direct reports
        /// </summary>
        /// TODO Edit XML Comment Template for UsersDirectReports
        UsersDirectReports,

        /// <summary>
        ///     The users groups
        /// </summary>
        /// TODO Edit XML Comment Template for UsersGroups
        UsersGroups,

        /// <summary>
        ///     The users summaries
        /// </summary>
        /// TODO Edit XML Comment Template for UsersSummaries
        UsersSummaries,

        /// <summary>
        ///     The ou computers
        /// </summary>
        /// TODO Edit XML Comment Template for OuComputers
        OuComputers,

        /// <summary>
        ///     The ou groups
        /// </summary>
        /// TODO Edit XML Comment Template for OuGroups
        OuGroups,

        /// <summary>
        ///     The ou groups users
        /// </summary>
        /// TODO Edit XML Comment Template for OuGroupsUsers
        OuGroupsUsers,

        /// <summary>
        ///     The ou users
        /// </summary>
        /// TODO Edit XML Comment Template for OuUsers
        OuUsers,

        /// <summary>
        ///     The ou users direct reports
        /// </summary>
        /// TODO Edit XML Comment Template for OuUsersDirectReports
        OuUsersDirectReports,

        /// <summary>
        ///     The ou users groups
        /// </summary>
        /// TODO Edit XML Comment Template for OuUsersGroups
        OuUsersGroups,

        /// <summary>
        ///     The search computer
        /// </summary>
        /// TODO Edit XML Comment Template for SearchComputer
        SearchComputer,

        /// <summary>
        ///     The search group
        /// </summary>
        /// TODO Edit XML Comment Template for SearchGroup
        SearchGroup,

        /// <summary>
        ///     The search user
        /// </summary>
        /// TODO Edit XML Comment Template for SearchUser
        SearchUser
    }

    /// <summary>
    ///     Class Query.
    /// </summary>
    /// TODO Edit XML Comment Template for Query
    public class Query
    {
        /// <summary>
        ///     The cancellation token source
        /// </summary>
        /// TODO Edit XML Comment Template for _cancellationTokenSource
        private CancellationTokenSource _cancellationTokenSource;

        /// <summary>
        ///     Initializes a new instance of the <see cref="Query" />
        ///     class.
        /// </summary>
        /// <param name="queryType">Type of the query.</param>
        /// <param name="scope">The scope.</param>
        /// <param name="distinguishedNames">The distinguished names.</param>
        /// <param name="searchText">The search text.</param>
        /// TODO Edit XML Comment Template for #ctor
        public Query(
            QueryType queryType,
            Scope scope = null,
            IList<string> distinguishedNames = null,
            string searchText = null)
        {
            QueryType = queryType;
            Scope = scope;
            DistinguishedNames = distinguishedNames;
            SearchText = searchText;
        }

        /// <summary>
        ///     Gets the type of the query.
        /// </summary>
        /// <value>The type of the query.</value>
        /// TODO Edit XML Comment Template for QueryType
        public QueryType QueryType
        {
            get;
        }

        /// <summary>
        ///     Gets the scope.
        /// </summary>
        /// <value>The scope.</value>
        /// TODO Edit XML Comment Template for Scope
        public Scope Scope
        {
            get;
        }

        /// <summary>
        ///     Gets the data.
        /// </summary>
        /// <value>The data.</value>
        /// TODO Edit XML Comment Template for Data
        public IEnumerable<ExpandoObject> Data
        {
            get;
            private set;
        }

        /// <summary>
        ///     Gets the cancellation token.
        /// </summary>
        /// <value>The cancellation token.</value>
        /// TODO Edit XML Comment Template for CancellationToken
        private CancellationToken CancellationToken => _cancellationTokenSource
            .Token;

        /// <summary>
        ///     Gets the distinguished names.
        /// </summary>
        /// <value>The distinguished names.</value>
        /// TODO Edit XML Comment Template for DistinguishedNames
        private IList<string> DistinguishedNames
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
        ///     Cancels this instance.
        /// </summary>
        /// TODO Edit XML Comment Template for Cancel
        public void Cancel()
        {
            _cancellationTokenSource?.Cancel();
        }

        /// <summary>
        ///     Disposes the data.
        /// </summary>
        /// TODO Edit XML Comment Template for DisposeData
        public void DisposeData()
        {
            Data = null;
        }

        /// <summary>
        ///     Executes this instance.
        /// </summary>
        /// <returns>Task.</returns>
        /// TODO Edit XML Comment Template for Execute
        public async Task Execute()
        {
            _cancellationTokenSource = new CancellationTokenSource();

            var task = Task.Run(
                () =>
                {
                    Data = new Searcher(
                        QueryType,
                        Scope,
                        DistinguishedNames,
                        CancellationToken,
                        SearchText).GetData();
                },
                _cancellationTokenSource.Token);

            await task;
        }

        /// <summary>
        ///     Returns a <see cref="System.String" /> that represents
        ///     this instance.
        /// </summary>
        /// <returns>
        ///     A <see cref="System.String" /> that represents
        ///     this instance.
        /// </returns>
        /// TODO Edit XML Comment Template for ToString
        public override string ToString()
        {
            return QueryType.ToString();
        }
    }
}