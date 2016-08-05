using System.Collections.Generic;
using System.Dynamic;
using System.Threading;
using System.Threading.Tasks;

namespace MichaelBrandonMorris.ActiveDirectoryTool
{
    public enum QueryType
    {
        None,
        ComputersGroups,
        ComputersSummaries,
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

    public class Query
    {
        private CancellationTokenSource _cancellationTokenSource;

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

        public IEnumerable<ExpandoObject> Data
        {
            get;
            private set;
        }

        public QueryType QueryType
        {
            get;
        }

        public Scope Scope
        {
            get;
        }

        private CancellationToken CancellationToken
            => _cancellationTokenSource.Token;

        private IList<string> DistinguishedNames
        {
            get;
        }

        private string SearchText
        {
            get;
        }

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

        public override string ToString()
        {
            return QueryType.ToString();
        }
    }
}