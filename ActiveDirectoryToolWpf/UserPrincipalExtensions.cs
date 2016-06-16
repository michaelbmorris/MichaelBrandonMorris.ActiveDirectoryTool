using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Linq;

namespace ActiveDirectoryToolWpf
{
    internal static class UserPrincipalExtensions
    {
        public const string Assistant = "assistant";
        public const string City = "l";
        public const string Comment = "comment";
        public const string Company = "company";
        public const string Country = "c";
        public const string Department = "department";
        public const string DirectReports = "directReports";
        public const string Division = "division";
        public const string Fax = "facsimileTelephoneNumber";
        public const string HomeAddress = "homePostalAddress";
        public const string HomePhone = "homePhone";
        public const string Initials = "initials";
        public const string Manager = "manager";
        public const string Mobile = "mobile";
        public const string None = "None";
        public const string Notes = "info";
        public const string Pager = "pager";
        public const string Sip = "msRTCSIP-PrimaryUserAddress";
        public const string State = "st";
        public const string StreetAddress = "streetAddress";
        public const string Suffix = "generationQualifier";
        public const string Title = "title";
        public const string UserAccountControl = "userAccountControl";
        public const string Voip = "ipPhone";

        internal static string GetAssistant(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Assistant);
        }

        internal static string GetCity(this UserPrincipal user)
        {
            return user.GetPropertyAsString(City);
        }

        internal static string GetComment(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Comment);
        }

        internal static string GetCompany(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Company);
        }

        internal static string GetCountry(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Country);
        }

        internal static string GetDepartment(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Department);
        }

        internal static IEnumerable<string> GetDirectReportDistinguishedNames(
            this UserPrincipal user)
        {
            return (IEnumerable<string>)user.GetProperty(DirectReports);
        }

        internal static IEnumerable<UserPrincipal> GetDirectReports(
            this UserPrincipal user)
        {
            return user.GetDirectReportDistinguishedNames()
                .Select(directReportDistinguishedName =>
                UserPrincipal.FindByIdentity(
                    user.Context,
                    IdentityType.DistinguishedName,
                    directReportDistinguishedName)).ToList();
        }

        internal static string GetDivision(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Division);
        }

        internal static string GetFax(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Fax);
        }

        internal static string GetHomeAddress(this UserPrincipal user)
        {
            return user.GetPropertyAsString(HomeAddress);
        }

        internal static string GetHomePhone(this UserPrincipal user)
        {
            return user.GetPropertyAsString(HomePhone);
        }

        internal static string GetInitials(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Initials);
        }

        internal static string GetManager(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Manager);
        }

        internal static string GetMobile(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Mobile);
        }

        internal static string GetNotes(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Notes);
        }

        internal static string GetPager(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Pager);
        }

        internal static string GetSip(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Sip);
        }

        internal static string GetState(this UserPrincipal user)
        {
            return user.GetPropertyAsString(State);
        }

        internal static string GetStreetAddress(this UserPrincipal user)
        {
            return user.GetPropertyAsString(StreetAddress);
        }

        internal static string GetSuffix(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Suffix);
        }

        internal static string GetTitle(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Title);
        }

        internal static string GetUserAccountControl(this UserPrincipal user)
        {
            return user.GetPropertyAsString(UserAccountControl);
        }

        internal static string GetVoip(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Voip);
        }

        internal static bool IsActive(this UserPrincipal user)
        {
            return !Convert.ToBoolean(
                (int)user.GetProperty(UserAccountControl) & 0x0002);
        }
    }
}