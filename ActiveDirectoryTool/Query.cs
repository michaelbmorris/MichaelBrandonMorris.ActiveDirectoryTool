using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Extensions.CollectionExtensions;
using static ActiveDirectoryTool.QueryType;
using static ActiveDirectoryTool.SimplifiedQueryType;
using static ActiveDirectoryTool.ActiveDirectoryProperty;

using DataMapping = System.Collections.Generic.Dictionary<ActiveDirectoryTool.QueryType, System.Func<System.Collections.Generic.IEnumerable<System.Dynamic.ExpandoObject>>>;

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
        private static readonly PrincipalContext RootContext = 
            new PrincipalContext(ContextType.Domain);

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
                Data = new DataPreparer(QueryType, _activeDirectoryScope, _distinguishedNames, CancellationToken).GetData();
            },
            _cancellationTokenSource.Token);
            await task;
        }
    }
}