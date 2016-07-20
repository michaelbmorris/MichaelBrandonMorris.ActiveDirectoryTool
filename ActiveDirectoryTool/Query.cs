using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Threading;
using System.Threading.Tasks;

namespace ActiveDirectoryTool
{
    public enum QueryType
    {
        None,
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
        OuUsersGroups,
        SearchComputer,
        SearchGroup,
        SearchUser
    }

    internal enum SimplifiedQueryType
    {
        DirectReports,
        Groups,
        Summaries
    }

    public class Query
    {
        private readonly Scope _activeDirectoryScope;
        private readonly IEnumerable<string> _distinguishedNames;
        private readonly string _searchText;

        private CancellationTokenSource _cancellationTokenSource;

        public Query(
            QueryType queryType,
            Scope activeDirectoryScope = null,
            IEnumerable<string> distinguishedNames = null,
            string searchText = null)
        {
            QueryType = queryType;
            _activeDirectoryScope = activeDirectoryScope;
            _distinguishedNames = distinguishedNames;
            _searchText = searchText;
        }

        private CancellationToken CancellationToken
            => _cancellationTokenSource.Token;

        public IEnumerable<ExpandoObject> Data { get; private set; }

        public string Name => Scope + " - " + QueryType;

        public QueryType QueryType { get; }

        public string Scope { get;  }

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

            var task = Task.Run(() =>
            {
                Data = new Searcher(
                    QueryType,
                    _activeDirectoryScope,
                    _distinguishedNames,
                    CancellationToken,
                    _searchText).GetData();
            },
            _cancellationTokenSource.Token);
            await task;
        }
    }
}