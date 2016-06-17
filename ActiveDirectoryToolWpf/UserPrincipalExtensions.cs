using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using PrimitiveExtensions;

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

        public static string GetAssistant(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Assistant);
        }

        public static string GetCity(this UserPrincipal user)
        {
            return user.GetPropertyAsString(City);
        }

        public static string GetComment(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Comment);
        }

        public static string GetCompany(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Company);
        }

        public static string GetCountry(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Country);
        }

        public static string GetDepartment(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Department);
        }

        public static IEnumerable<UserPrincipal> GetDirectReports(
            this UserPrincipal user)
        {
            var directReportDistinguishedNames =
                user.GetDirectReportDistinguishedNames();
            if (directReportDistinguishedNames != null)
                return user.GetDirectReportDistinguishedNames()
                    .Select(directReportDistinguishedName =>
                        UserPrincipal.FindByIdentity(
                            user.Context,
                            IdentityType.DistinguishedName,
                            directReportDistinguishedName)).ToList();
            return null;
        }

        public static string GetDivision(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Division);
        }

        public static string GetFax(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Fax);
        }

        public static string GetHomeAddress(this UserPrincipal user)
        {
            return user.GetPropertyAsString(HomeAddress);
        }

        public static string GetHomePhone(this UserPrincipal user)
        {
            return user.GetPropertyAsString(HomePhone);
        }

        public static string GetInitials(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Initials);
        }

        public static string GetManager(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Manager);
        }

        public static string GetMobile(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Mobile);
        }

        public static string GetNotes(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Notes);
        }

        public static string GetPager(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Pager);
        }

        public static string GetSip(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Sip);
        }

        public static string GetState(this UserPrincipal user)
        {
            return user.GetPropertyAsString(State);
        }

        public static string GetStreetAddress(this UserPrincipal user)
        {
            return user.GetPropertyAsString(StreetAddress);
        }

        public static string GetSuffix(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Suffix);
        }

        public static string GetTitle(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Title);
        }

        public static string GetUserAccountControl(this UserPrincipal user)
        {
            return user.GetPropertyAsString(UserAccountControl);
        }

        public static string GetVoip(this UserPrincipal user)
        {
            return user.GetPropertyAsString(Voip);
        }

        public static bool IsActive(this UserPrincipal user)
        {
            return !Convert.ToBoolean(
                (int) user.GetProperty(UserAccountControl).Value & 0x0002);
        }

        internal static IEnumerable<string> GetDirectReportDistinguishedNames(
            this UserPrincipal user)
        {
            return
                user.GetProperty(DirectReports)
                    .Cast<string>()
                    .Where(
                        directReportDistinguishedName =>
                            !directReportDistinguishedName
                                .IsNullOrWhiteSpace())
                    .ToList();
        }
    }
}