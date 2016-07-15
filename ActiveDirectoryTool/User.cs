using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.Security.Cryptography.X509Certificates;
using Extensions.PrincipalExtensions;

namespace ActiveDirectoryTool
{
    public interface IAuthenticablePrincipal : IPrincipal
    {
        DateTime? AccountExpirationDate { get; set; }
        DateTime? AccountLockoutTime { get; set; }
        bool AllowReversiblePasswordEncryption { get; set; }
        int BadLogonCount { get; set; }
        X509Certificate2Collection Certificates { get; set; }
        bool? DelegationPermitted { get; set; }
        bool? Enabled { get; set; }
        string HomeDirectory { get; set; }
        string HomeDrive { get; set; }
        DateTime? LastBadPasswordAttempt { get; set; }
        DateTime? LastLogon { get; set; }
        DateTime? LastPasswordSet { get; set; }
        bool PasswordNeverExpires { get; set; }
        bool PasswordNotRequired { get; set; }
        byte[] PermittedLogonTimes { get; set; }
        PrincipalValueCollection<string> PermittedWorkstations { get; set; }
        string ScriptPath { get; set; }
        bool SmartCardLogonRequired { get; set; }
    }

    public interface IComputerPrincipal : IAuthenticablePrincipal
    {
    }

    public interface IGroupPrincipal : IPrincipal
    {
    }

    public interface IPrincipal
    {
        PrincipalContext Context { get; set; }
        ContextType ContextType { get; set; }
        string Description { get; set; }
        string DisplayName { get; set; }
        string DistinguishedName { get; set; }
        Guid? Guid { get; set; }
        string Name { get; set; }
        string SamAccountName { get; set; }
        string StructuralObjectClass { get; set; }
        string UserPrincipalName { get; set; }
    }

    public interface IUserDirectoryEntry
    {
        string Assistant { get; set; }
        string City { get; set; }
        string Comment { get; set; }
        string Company { get; set; }
        string Country { get; set; }
        string Department { get; set; }
        IEnumerable<User> DirectReports { get; set; }
        string Division { get; set; }
        string Fax { get; set; }
        string HomeAddress { get; set; }
        string HomePhone { get; set; }
        string Initials { get; set; }
        bool IsActive { get; set; }
        User Manager { get; set; }
        string Mobile { get; set; }
        string Notes { get; set; }
        string Pager { get; set; }
        string Sip { get; set; }
        string State { get; set; }
        string StreetAddress { get; set; }
        string Suffix { get; set; }
        string Title { get; set; }
        string Voip { get; set; }
    }

    public interface IUserPrincipal : IAuthenticablePrincipal
    {
        string EmailAddress { get; set; }
        string EmployeeId { get; set; }
        string GivenName { get; set; }
        string MiddleName { get; set; }
        string Surname { get; set; }
        bool UserCannotChangePassword { get; set; }
        string VoiceTelephoneNumber { get; set; }
    }

    public class AuthenticablePrincipalWrapper : PrincipalWrapper,
        IAuthenticablePrincipal
    {
        protected AuthenticablePrincipalWrapper(
            AuthenticablePrincipal authenticablePrincipal)
            : base(authenticablePrincipal)
        {
            AccountExpirationDate =
                authenticablePrincipal.AccountExpirationDate;

            AccountLockoutTime = authenticablePrincipal.AccountLockoutTime;

            AllowReversiblePasswordEncryption =
                authenticablePrincipal.AllowReversiblePasswordEncryption;

            BadLogonCount = authenticablePrincipal.BadLogonCount;
            Certificates = authenticablePrincipal.Certificates;
            DelegationPermitted = authenticablePrincipal.DelegationPermitted;
            Enabled = authenticablePrincipal.Enabled;
        }

        public DateTime? AccountExpirationDate { get; set; }
        public DateTime? AccountLockoutTime { get; set; }
        public bool AllowReversiblePasswordEncryption { get; set; }
        public int BadLogonCount { get; set; }
        public X509Certificate2Collection Certificates { get; set; }
        public bool? DelegationPermitted { get; set; }
        public bool? Enabled { get; set; }
        public string HomeDirectory { get; set; }
        public string HomeDrive { get; set; }
        public DateTime? LastBadPasswordAttempt { get; set; }
        public DateTime? LastLogon { get; set; }
        public DateTime? LastPasswordSet { get; set; }
        public bool PasswordNeverExpires { get; set; }
        public bool PasswordNotRequired { get; set; }
        public byte[] PermittedLogonTimes { get; set; }

        public PrincipalValueCollection<string> PermittedWorkstations { get;
            set; }

        public string ScriptPath { get; set; }
        public bool SmartCardLogonRequired { get; set; }
    }

    public class PrincipalWrapper : IPrincipal
    {
        protected PrincipalWrapper(Principal principal)
        {
            Context = principal.Context;
            ContextType = principal.ContextType;
            Description = principal.Description;
            DisplayName = principal.DisplayName;
            DistinguishedName = principal.DistinguishedName;
            Guid = principal.Guid;
            Name = principal.Name;
            SamAccountName = principal.SamAccountName;
            StructuralObjectClass = principal.StructuralObjectClass;
            UserPrincipalName = principal.UserPrincipalName;
            Debug.WriteLine(Name);
        }

        public PrincipalContext Context { get; set; }
        public ContextType ContextType { get; set; }
        public string Description { get; set; }
        public string DisplayName { get; set; }
        public string DistinguishedName { get; set; }
        public Guid? Guid { get; set; }
        public string Name { get; set; }
        public string SamAccountName { get; set; }
        public string StructuralObjectClass { get; set; }
        public string UserPrincipalName { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }

    public class User : AuthenticablePrincipalWrapper, IUserPrincipal,
        IUserDirectoryEntry
    {
        private static readonly PrincipalContext PrincipalContext =
            new PrincipalContext(ContextType.Domain);

        public User(UserPrincipal userPrincipal)
            : base(userPrincipal)
        {
            EmailAddress = userPrincipal.EmailAddress;
            EmployeeId = userPrincipal.EmployeeId;
            GivenName = userPrincipal.GivenName;
            MiddleName = userPrincipal.MiddleName;
            Surname = userPrincipal.Surname;
            VoiceTelephoneNumber = userPrincipal.VoiceTelephoneNumber;
            UserCannotChangePassword = userPrincipal.UserCannotChangePassword;
            Assistant = userPrincipal.GetAssistant();
            City = userPrincipal.GetCity();
            Comment = userPrincipal.GetComment();
            Company = userPrincipal.GetCompany();
            Country = userPrincipal.GetCountry();
            Department = userPrincipal.GetDepartment();
            //DirectReports = GetDirectReports(userPrincipal);
            Division = userPrincipal.GetDivision();
            Fax = userPrincipal.GetFax();
            HomeAddress = userPrincipal.GetHomeAddress();
            HomePhone = userPrincipal.GetHomePhone();
            Initials = userPrincipal.GetInitials();
            IsActive = userPrincipal.IsActive();
            //Manager = GetManager(userPrincipal);
            Mobile = userPrincipal.GetMobile();
            Notes = userPrincipal.GetNotes();
            Pager = userPrincipal.GetPager();
            Sip = userPrincipal.GetSip();
            State = userPrincipal.GetState();
            StreetAddress = userPrincipal.GetStreetAddress();
            Suffix = userPrincipal.GetSuffix();
            Title = userPrincipal.GetTitle();
            Voip = userPrincipal.GetVoip();
        }

        public string Assistant { get; set; }
        public string City { get; set; }
        public string Comment { get; set; }
        public string Company { get; set; }
        public string Country { get; set; }
        public string Department { get; set; }
        public IEnumerable<User> DirectReports { get; set; }
        public string Division { get; set; }
        public string EmailAddress { get; set; }
        public string EmployeeId { get; set; }
        public string Fax { get; set; }
        public string GivenName { get; set; }
        public string HomeAddress { get; set; }
        public string HomePhone { get; set; }
        public string Initials { get; set; }
        public bool IsActive { get; set; }
        public User Manager { get; set; }
        public string MiddleName { get; set; }
        public string Mobile { get; set; }
        public string Notes { get; set; }
        public string Pager { get; set; }
        public string Sip { get; set; }
        public string State { get; set; }
        public string StreetAddress { get; set; }
        public string Suffix { get; set; }
        public string Surname { get; set; }
        public string Title { get; set; }
        public bool UserCannotChangePassword { get; set; }
        public string VoiceTelephoneNumber { get; set; }
        public string Voip { get; set; }

        private static UserPrincipal FindUser(string distinguishedName)
        {
            return UserPrincipal.FindByIdentity(
                PrincipalContext, IdentityType.DistinguishedName,
                distinguishedName);
        }

        private static IEnumerable<User> GetDirectReports(
            UserPrincipal userPrincipal)
        {
            var directReports = new List<User>();
            foreach (var directReportDistinguishedName in 
                userPrincipal.GetDirectReportDistinguishedNames())
            {
                using (var directReportUserPrincipal =
                    FindUser(directReportDistinguishedName))
                {
                    directReports.Add(new User(directReportUserPrincipal));
                }
            }
            return directReports;
        }

        private static User GetManager(UserPrincipal userPrincipal)
        {
            using (var managerUserPrincipal =
                FindUser(userPrincipal.GetManager()))
            {
                return new User(managerUserPrincipal);
            }
        }
    }
}