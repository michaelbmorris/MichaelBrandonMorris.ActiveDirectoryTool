using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiveDirectoryToolWpf
{
    public static class PrincipalSearcherExtensions
    {
        public static IEnumerable<ComputerPrincipal> GetAllComputerPrincipals(
            this PrincipalSearcher principalSearcher)
        {
            return principalSearcher.FindAll().GetComputerPrincipals();
        }

        public static IEnumerable<UserPrincipal> GetAllUserPrincipals(
            this PrincipalSearcher principalSearcher)
        {
            return principalSearcher.FindAll().GetUserPrincipals();
        }
    }
}
