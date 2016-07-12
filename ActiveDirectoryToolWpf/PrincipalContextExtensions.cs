using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiveDirectoryToolWpf
{
    public static class PrincipalContextExtensions
    {
        public static Principal FindByDistinguishedName(
            this PrincipalContext principalContext, string distinguishedName)
        {
            return Principal.FindByIdentity(
                principalContext,
                IdentityType.DistinguishedName,
                distinguishedName);
        }

        public static UserPrincipal FindUserByDistinguishedName(
            this PrincipalContext principalContext, string distinguishedName)
        {
            return UserPrincipal.FindByIdentity(
                principalContext,
                IdentityType.DistinguishedName,
                distinguishedName);
        }

        public static IEnumerable<UserPrincipal> FindUsersByDistinguishedNames(
            this PrincipalContext principalContext,
            IEnumerable<string> distinguishedNames)
        {
            return distinguishedNames.Select(
                principalContext.FindUserByDistinguishedName);
        }

        public static IEnumerable<Principal> FindByDistinguishedNames(
            this PrincipalContext principalContext,
            IEnumerable<string> distinguishedNames)
        {
            return distinguishedNames.Select(
                principalContext.FindByDistinguishedName);
        }
    }
}
