using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Linq;

namespace ActiveDirectoryToolWpf
{
    public static class PrincipalSearchResultExtensions
    {
        public static IEnumerable<ComputerPrincipal> GetComputerPrincipals(
            this PrincipalSearchResult<Principal> principalSearchResult)
        {
            return principalSearchResult.OfType<ComputerPrincipal>();
        }

        public static IEnumerable<GroupPrincipal> GetGroupPrincipals(
            this PrincipalSearchResult<Principal> principalSearchResult)
        {
            return principalSearchResult.OfType<GroupPrincipal>();
        }

        public static IEnumerable<UserPrincipal> GetUserPrincipals(
            this PrincipalSearchResult<Principal> principalSearchResult)
        {
            return principalSearchResult.OfType<UserPrincipal>();
        }
    }
}