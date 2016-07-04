using System;
using System.Collections.Generic;
using System.Data;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Linq;
using static ActiveDirectoryToolWpf.ActiveDirectoryAttribute;

namespace ActiveDirectoryToolWpf
{
    public enum ActiveDirectoryAttribute
    {
        ComputerAccountExpirationDate,
        ComputerAccountLockoutTime,
        ComputerAllowReversiblePasswordEncryption,
        ComputerBadLogonCount,
        ComputerCertificates,
        ComputerContext,
        ComputerContextType,
        ComputerDelegationPermitted,
        ComputerDescription,
        ComputerDisplayName,
        ComputerDistinguishedName,
        ComputerEnabled,
        ComputerGuid,
        ComputerHomeDirectory,
        ComputerHomeDrive,
        ComputerLastBadPasswordAttempt,
        ComputerLastLogon,
        ComputerLastPasswordSet,
        ComputerName,
        ComputerPasswordNeverExpires,
        ComputerPasswordNotRequired,
        ComputerPermittedLogonTimes,
        ComputerPermittedWorkstations,
        ComputerSamAccountName,
        ComputerScriptPath,
        ComputerServicePrincipalNames,
        ComputerSid,
        ComputerSmartcardLogonRequired,
        ComputerStructuralObjectClass,
        ComputerUserCannotChangePassword,
        ComputerUserPrincipalName,
        DirectReportUserAccountControl,
        DirectReportAccountExpirationDate,
        DirectReportAccountLockoutTime,
        DirectReportAllowReversiblePasswordEncryption,
        DirectReportAssistant,
        DirectReportBadLogonCount,
        DirectReportCertificates,
        DirectReportCity,
        DirectReportComment,
        DirectReportCompany,
        DirectReportContext,
        DirectReportContextType,
        DirectReportCountry,
        DirectReportDelegationPermitted,
        DirectReportDepartment,
        DirectReportDescription,
        DirectReportDisplayName,
        DirectReportDistinguishedName,
        DirectReportDivision,
        DirectReportEmailAddress,
        DirectReportEmployeeId,
        DirectReportEnabled,
        DirectReportFax,
        DirectReportSuffix,
        DirectReportGivenName,
        DirectReportGuid,
        DirectReportHomeAddress,
        DirectReportHomeDirectory,
        DirectReportHomeDrive,
        DirectReportHomePhone,
        DirectReportInitials,
        DirectReportIsAccountLockedOut,
        DirectReportIsActive,
        DirectReportLastBadPasswordAttempt,
        DirectReportLastLogon,
        DirectReportLastPasswordSet,
        DirectReportManager,
        DirectReportMiddleName,
        DirectReportMobile,
        DirectReportName,
        DirectReportNotes,
        DirectReportPager,
        DirectReportPasswordNeverExpires,
        DirectReportPasswordNotRequired,
        DirectReportPermittedLogonTimes,
        DirectReportPermittedWorkstations,
        DirectReportSamAccountName,
        DirectReportScriptPath,
        DirectReportSid,
        DirectReportSip,
        DirectReportSmartcardLogonRequired,
        DirectReportState,
        DirectReportStreetAddress,
        DirectReportStructuralObjectClass,
        DirectReportSurname,
        DirectReportTitle,
        DirectReportUserCannotChangePassword,
        DirectReportUserPrincipalName,
        DirectReportVoiceTelephoneNumber,
        DirectReportVoip,
        GroupContext,
        GroupContextType,
        GroupDescription,
        GroupDisplayName,
        GroupDistinguishedName,
        GroupGuid,
        GroupIsSecurityGroup,
        GroupManagedBy,
        GroupMembers,
        GroupName,
        GroupSamAccountName,
        GroupScope,
        GroupSid,
        GroupStructuralObjectClass,
        GroupUserPrincipalName,
        UserUserAccountControl,
        UserAccountExpirationDate,
        UserAccountLockoutTime,
        UserAllowReversiblePasswordEncryption,
        UserAssistant,
        UserBadLogonCount,
        UserCertificates,
        UserCity,
        UserComment,
        UserCompany,
        UserContext,
        UserContextType,
        UserCountry,
        UserDelegationPermitted,
        UserDepartment,
        UserDescription,
        UserDisplayName,
        UserDistinguishedName,
        UserDivision,
        UserEmailAddress,
        UserEmployeeId,
        UserEnabled,
        UserFax,
        UserSuffix,
        UserGivenName,
        UserGuid,
        UserHomeAddress,
        UserHomeDirectory,
        UserHomeDrive,
        UserHomePhone,
        UserInitials,
        UserIsAccountLockedOut,
        UserIsActive,
        UserLastBadPasswordAttempt,
        UserLastLogon,
        UserLastPasswordSet,
        UserManager,
        UserMiddleName,
        UserMobile,
        UserName,
        UserNotes,
        UserPager,
        UserPasswordNeverExpires,
        UserPasswordNotRequired,
        UserPermittedLogonTimes,
        UserPermittedWorkstations,
        UserSamAccountName,
        UserScriptPath,
        UserSid,
        UserSip,
        UserSmartcardLogonRequired,
        UserState,
        UserStreetAddress,
        UserStructuralObjectClass,
        UserSurname,
        UserTitle,
        UserUserCannotChangePassword,
        UserUserPrincipalName,
        UserVoiceTelephoneNumber,
        UserVoip
    }

    internal static class Extensions
    {
        public static DataTable ToDataTable(this IEnumerable<dynamic> items)
        {
            var data = items.ToArray();
            if (data.Length == 0) return null;
            var dataTable = new DataTable();
            foreach (
                var key in ((IDictionary<string, object>) data[0]).Keys)
            {
                dataTable.Columns.Add(key);
            }
            foreach (var d in data)
            {
                dataTable.Rows.Add(
                    ((IDictionary<string, object>) d).Values.ToArray());
            }
            return dataTable;
        }
    }

    internal class DataPreparer
    {
        internal List<ActiveDirectoryAttribute> Attributes { get; set; }
        internal IEnumerable<object> Data { get; set; }

        internal IEnumerable<ExpandoObject> GetResults()
        {
            var results = new List<ExpandoObject>();

            foreach (var data in Data)
            {
                if (data is UserGroups)
                {
                    var userGroups = data as UserGroups;
                    foreach (var group in userGroups.Groups)
                    {
                        dynamic result = new ExpandoObject();
                        AddAttributesToResult(
                            result,
                            groupPrincipal:group,
                            userPrincipal:userGroups.User);
                        results.Add(result);
                    }
                    userGroups.Dispose();
                }
                else if (data is UserDirectReports)
                {
                    var directReports = data as UserDirectReports;
                    if (directReports.DirectReports == null) continue;
                    foreach (var directReport in directReports.DirectReports)
                    {
                        dynamic result = new ExpandoObject();
                        AddAttributesToResult(
                            result,
                            directReportUserPrincipal:directReport,
                            userPrincipal:directReports.User);
                        results.Add(result);
                    }
                    directReports.Dispose();
                }
                else if (data is ComputerGroups)
                {
                    var computerGroups = data as ComputerGroups;
                    if (computerGroups.Groups == null) continue;
                    foreach (var group in computerGroups.Groups)
                    {
                        dynamic result = new ExpandoObject();
                        AddAttributesToResult(
                            result,
                            computerPrincipal:computerGroups.Computer,
                            groupPrincipal:group);
                        results.Add(result);
                    }
                }
                else
                {
                    var computerPrincipal = data as ComputerPrincipal;
                    var groupPrincipal = data as GroupPrincipal;
                    var userPrincipal = data as UserPrincipal;
                    dynamic result = new ExpandoObject();
                    AddAttributesToResult(
                        result,
                        computerPrincipal,
                        groupPrincipal,
                        null,
                        userPrincipal);
                    results.Add(result);
                    var principal = data as Principal;
                    principal?.Dispose();
                }
            }
            return results;
        }

        private static bool AddComputerAttributeToResult(
            ComputerPrincipal principal,
            ActiveDirectoryAttribute attribute,
            dynamic result)
        {
            var attributeMapping =
                new Dictionary<ActiveDirectoryAttribute, Action>
                {
                    [ComputerAccountExpirationDate] = () =>
                        result.ComputerAccountExpirationDate =
                            principal.AccountExpirationDate,
                    [ComputerAccountLockoutTime] = () =>
                        result.ComputerAccountLockoutTime =
                            principal.AccountLockoutTime,
                    [ComputerAllowReversiblePasswordEncryption] = () =>
                        result.ComputerAllowReversiblePasswordEncryption =
                            principal.AllowReversiblePasswordEncryption,
                    [ComputerBadLogonCount] = () =>
                        result.ComputerBadLogonCount = principal.BadLogonCount,
                    [ComputerCertificates] = () =>
                        result.ComputerCertificates = principal.Certificates,
                    [ComputerContext] = () =>
                        result.ComputerContext = principal.Context,
                    [ComputerContextType] = () =>
                        result.ComputerContextType = principal.ContextType,
                    [ComputerDelegationPermitted] = () =>
                        result.ComputerDelegationPermitted =
                            principal.DelegationPermitted,
                    [ComputerDescription] = () =>
                        result.ComputerDescription = principal.Description,
                    [ComputerDisplayName] = () =>
                        result.ComputerDisplayName = principal.DisplayName,
                    [ComputerDistinguishedName] = () =>
                        result.ComputerDistinguishedName =
                            principal.DistinguishedName,
                    [ComputerEnabled] = () =>
                        result.ComputerEnabled = principal.Enabled,
                    [ComputerGuid] = () => result.ComputerGuid =
                        principal.Guid,
                    [ComputerHomeDirectory] = () =>
                        result.ComputerHomeDirectory = principal.HomeDirectory,
                    [ComputerHomeDrive] = () =>
                        result.ComputerHomeDrive = principal.HomeDrive,
                    [ComputerLastBadPasswordAttempt] = () =>
                        result.ComputerLastBadPasswordAttempt =
                            principal.LastBadPasswordAttempt,
                    [ComputerLastLogon] = () =>
                        result.ComputerLastLogon = principal.LastLogon,
                    [ComputerLastPasswordSet] = () =>
                        result.LastPasswordSet = principal.LastPasswordSet,
                    [ComputerName] = () => result.ComputerName =
                        principal.Name,
                    [ComputerPasswordNeverExpires] = () =>
                        result.ComputerPasswordNeverExpires =
                            principal.PasswordNeverExpires,
                    [ComputerPasswordNotRequired] = () =>
                        result.ComputerPasswordNotRequired =
                            principal.PasswordNotRequired,
                    [ComputerPermittedLogonTimes] = () =>
                        result.ComputerPermittedLogonTimes =
                            principal.PermittedLogonTimes,
                    [ComputerPermittedWorkstations] = () =>
                        result.ComputerPermittedWorkstations =
                            principal.PermittedWorkstations,
                    [ComputerSamAccountName] = () =>
                        result.ComputerSamAccountName =
                            principal.SamAccountName,
                    [ComputerScriptPath] = () =>
                        result.ComputerScriptPath = principal.ScriptPath,
                    [ComputerServicePrincipalNames] = () =>
                        result.ComputerServicePrincipalNames =
                            principal.ServicePrincipalNames,
                    [ComputerSid] = () => result.ComputerSid = principal.Sid,
                    [ComputerSmartcardLogonRequired] = () =>
                        result.ComputerSmartcardLogonRequired =
                            principal.SmartcardLogonRequired,
                    [ComputerStructuralObjectClass] = () =>
                        result.ComputerStructuralObjectClass =
                            principal.StructuralObjectClass,
                    [ComputerUserCannotChangePassword] = () =>
                        result.ComputerUserCannotChangePassword =
                            principal.UserCannotChangePassword,
                    [ComputerUserPrincipalName] = () =>
                        result.ComputerUserPrincipalName =
                            principal.UserPrincipalName
                };

            if (!attributeMapping.ContainsKey(attribute)) return false;
            attributeMapping[attribute]();
            return true;
        }

        private static bool AddDirectReportAttributeToResult(
            UserPrincipal principal,
            ActiveDirectoryAttribute attribute,
            dynamic result)
        {
            var attributeMapping =
                new Dictionary<ActiveDirectoryAttribute, Action>
                {
                    [DirectReportUserAccountControl] = () =>
                        result.DirectReportAccountControl =
                            principal.GetUserAccountControl(),
                    [DirectReportAccountExpirationDate] = () =>
                        result.DirectReportAccountExpirationDate =
                            principal.AccountExpirationDate,
                    [DirectReportAccountLockoutTime] = () =>
                        result.DirectReportAccountLockoutTime =
                            principal.AccountLockoutTime,
                    [DirectReportAllowReversiblePasswordEncryption] = () =>
                        result.DirectReportAllowReversiblePasswordEncryption =
                            principal.AllowReversiblePasswordEncryption,
                    [DirectReportAssistant] = () =>
                        result.DirectReportAssistant =
                            principal.GetAssistant(),
                    [DirectReportBadLogonCount] = () =>
                        result.DirectReportBadLogonCount =
                            principal.BadLogonCount,
                    [DirectReportCertificates] = () =>
                        result.DirectReportCertificates =
                            principal.Certificates,
                    [DirectReportCity] = () =>
                        result.DirectReportCity = principal.GetCity(),
                    [DirectReportComment] = () =>
                        result.DirectReportComment = principal.GetComment(),
                    [DirectReportCompany] = () =>
                        result.DirectReportCompany = principal.GetCompany(),
                    [DirectReportContext] = () =>
                        result.DirectReportContext = principal.Context,
                    [DirectReportContextType] = () =>
                        result.DirectReportContextType = principal.ContextType,
                    [DirectReportCountry] = () =>
                        result.DirectReportCountry = principal.GetCountry(),
                    [DirectReportDelegationPermitted] = () =>
                        result.DirectReportDelegationPermitted =
                            principal.DelegationPermitted,
                    [DirectReportDepartment] = () =>
                        result.DirectReportDepartment =
                            principal.GetDepartment(),
                    [DirectReportDescription] = () =>
                        result.DirectReportDescription = principal.Description,
                    [DirectReportDisplayName] = () =>
                        result.DirectReportDisplayName = principal.DisplayName,
                    [DirectReportDistinguishedName] = () =>
                        result.DirectReportDistinguishedName =
                            principal.DistinguishedName,
                    [DirectReportDivision] = () =>
                        result.DirectReportDivision = principal.GetDivision(),
                    [DirectReportEmailAddress] = () =>
                        result.DirectReportEmailAddress =
                            principal.EmailAddress,
                    [DirectReportEmployeeId] = () =>
                        result.DirectReportEmployeeId = principal.EmployeeId,
                    [DirectReportEnabled] = () =>
                        result.DirectReportEnabled = principal.Enabled,
                    [DirectReportFax] = () =>
                        result.DirectReportFax = principal.GetFax(),
                    [DirectReportSuffix] = () =>
                        result.DirectReportSuffix = principal.GetSuffix(),
                    [DirectReportGivenName] = () =>
                        result.DirectReportGivenName = principal.GivenName,
                    [DirectReportGuid] = () =>
                        result.DirectReportGuid = principal.Guid,
                    [DirectReportHomeAddress] = () =>
                        result.DirectReportHomeAddress =
                            principal.GetHomeAddress(),
                    [DirectReportHomeDirectory] = () =>
                        result.DirectReportHomeDirectory =
                            principal.HomeDirectory,
                    [DirectReportHomeDrive] = () =>
                        result.DirectReportHomeDrive = principal.HomeDrive,
                    [DirectReportHomePhone] = () =>
                        result.DirectReportHomePhone =
                            principal.GetHomePhone(),
                    [DirectReportInitials] = () =>
                        result.DirectReportInitials = principal.GetInitials(),
                    [DirectReportIsAccountLockedOut] = () =>
                        result.DirectReportIsAccountLockedOut =
                            principal.IsAccountLockedOut(),
                    [DirectReportIsActive] = () =>
                        result.DirectReportIsActive = principal.IsActive(),
                    [DirectReportLastBadPasswordAttempt] = () =>
                        result.DirectReportLastBadPasswordAttempt =
                            principal.LastBadPasswordAttempt,
                    [DirectReportLastLogon] = () =>
                        result.DirectReportLastLogon = principal.LastLogon,
                    [DirectReportLastPasswordSet] = () =>
                        result.DirectReportLastPasswordSet =
                            principal.LastPasswordSet,
                    [DirectReportManager] = () =>
                        result.DirectReportManager = principal.GetManager(),
                    [DirectReportMiddleName] = () =>
                        result.DirectReportMiddleName = principal.MiddleName,
                    [DirectReportMobile] = () =>
                        result.DirectReportMobile = principal.GetMobile(),
                    [DirectReportName] = () =>
                        result.DirectReportName = principal.Name,
                    [DirectReportNotes] = () =>
                        result.DirectReportNotes = principal.GetNotes(),
                    [DirectReportPager] = () =>
                        result.DirectReportPager = principal.GetPager(),
                    [DirectReportPasswordNeverExpires] = () =>
                        result.DirectReportPasswordNeverExpires =
                            principal.PasswordNeverExpires,
                    [DirectReportPasswordNotRequired] = () =>
                        result.DirectReportPasswordNotRequired =
                            principal.PasswordNotRequired,
                    [DirectReportPermittedLogonTimes] = () =>
                        result.DirectReportPermittedLogonTimes =
                            principal.PermittedLogonTimes,
                    [DirectReportPermittedWorkstations] = () =>
                        result.DirectReportPermittedWorkstations =
                            principal.PermittedWorkstations,
                    [DirectReportSamAccountName] = () =>
                        result.DirectReportSamAccountName =
                            principal.SamAccountName,
                    [DirectReportScriptPath] = () =>
                        result.DirectReportScriptPath = principal.ScriptPath,
                    [DirectReportSid] = () =>
                        result.DirectReportSid = principal.Sid,
                    [DirectReportSip] = () =>
                        result.DirectReportSip = principal.GetSip(),
                    [DirectReportSmartcardLogonRequired] = () =>
                        result.DirectReportSmartcardLogonRequired =
                            principal.SmartcardLogonRequired,
                    [DirectReportState] = () =>
                        result.DirectReportState = principal.GetState(),
                    [DirectReportStreetAddress] = () =>
                        result.DirectReportStreetAddress =
                            principal.GetStreetAddress(),
                    [DirectReportStructuralObjectClass] = () =>
                        result.DirectReportStructuralObjectClass =
                            principal.StructuralObjectClass,
                    [DirectReportSurname] = () =>
                        result.DirectReportSurname = principal.Surname,
                    [DirectReportTitle] = () =>
                        result.DirectReportTitle = principal.GetTitle(),
                    [DirectReportUserCannotChangePassword] = () =>
                        result.DirectReportUserCannotChangePassword =
                            principal.UserCannotChangePassword,
                    [DirectReportUserPrincipalName] = () =>
                        result.DirectReportUserPrincipalName =
                            principal.UserPrincipalName,
                    [DirectReportVoiceTelephoneNumber] = () =>
                        result.DirectReportVoiceTelephoneNumber =
                            principal.VoiceTelephoneNumber,
                    [DirectReportVoip] = () =>
                        result.DirectReportVoip = principal.GetVoip()
                };

            if (!attributeMapping.ContainsKey(attribute)) return false;
            attributeMapping[attribute]();
            return true;
        }

        private static bool AddGroupAttributeToResult(
            GroupPrincipal principal,
            ActiveDirectoryAttribute attribute,
            dynamic result)
        {
            var attributeMapping =
                new Dictionary<ActiveDirectoryAttribute, Action>
                {
                    [GroupContext] =
                        () => result.GroupContext = principal.Context,
                    [GroupContextType] = () =>
                        result.GroupContextType = principal.ContextType,
                    [GroupDescription] = () =>
                        result.GroupDescription = principal.Description,
                    [GroupDisplayName] = () =>
                        result.GroupDisplayName = principal.DisplayName,
                    [GroupDistinguishedName] = () =>
                        result.GroupDistinguishedName =
                            principal.DistinguishedName,
                    [GroupGuid] = () => result.GroupGuid = principal.Guid,
                    [GroupIsSecurityGroup] = () =>
                        result.GroupIsSecurityGroup =
                            principal.IsSecurityGroup,
                    [GroupManagedBy] = () =>
                        result.GroupManagedBy = principal.GetManagedBy(),
                    [GroupName] = () => result.GroupName = principal.Name,
                    [GroupSamAccountName] = () =>
                        result.GroupSamAccountName = principal.SamAccountName,
                    [ActiveDirectoryAttribute.GroupScope] = () =>
                        result.GroupScope = principal.GroupScope,
                    [GroupSid] = () => result.GroupSid = principal.Sid,
                    [GroupStructuralObjectClass] = () =>
                        result.GroupStructuralObjectClass =
                            principal.StructuralObjectClass,
                    [GroupUserPrincipalName] = () =>
                        result.GroupUserPrincipalName =
                            principal.UserPrincipalName,
                    [GroupMembers] =
                        () => result.GroupMembers = principal.Members
                };

            if (!attributeMapping.ContainsKey(attribute)) return false;
            attributeMapping[attribute]();
            return true;
        }

        private static void AddUserAttributeToResult(
            UserPrincipal principal,
            ActiveDirectoryAttribute attribute,
            dynamic result)
        {
            var attributeMapping =
                new Dictionary<ActiveDirectoryAttribute, Action>
                {
                    [UserUserAccountControl] = () =>
                        result.UserAccountControl =
                            principal.GetUserAccountControl(),
                    [UserAccountExpirationDate] = () =>
                        result.UserAccountExpirationDate =
                            principal.AccountExpirationDate,
                    [UserAccountLockoutTime] = () =>
                        result.UserAccountLockoutTime =
                            principal.AccountLockoutTime,
                    [UserAllowReversiblePasswordEncryption] = () =>
                        result.UserAllowReversiblePasswordEncryption =
                            principal.AllowReversiblePasswordEncryption,
                    [UserAssistant] = () =>
                        result.UserAssistant = principal.GetAssistant(),
                    [UserBadLogonCount] = () =>
                        result.UserBadLogonCount = principal.BadLogonCount,
                    [UserCertificates] = () =>
                        result.UserCertificates = principal.Certificates,
                    [UserCity] = () =>
                        result.UserCity = principal.GetCity(),
                    [UserComment] = () =>
                        result.UserComment = principal.GetComment(),
                    [UserCompany] = () =>
                        result.UserCompany = principal.GetCompany(),
                    [UserContext] = () =>
                        result.UserContext = principal.Context,
                    [UserContextType] = () =>
                        result.UserContextType = principal.ContextType,
                    [UserCountry] = () =>
                        result.UserCountry = principal.GetCountry(),
                    [UserDelegationPermitted] = () =>
                        result.UserDelegationPermitted =
                            principal.DelegationPermitted,
                    [UserDepartment] = () =>
                        result.UserDepartment = principal.GetDepartment(),
                    [UserDescription] = () =>
                        result.UserDescription = principal.Description,
                    [UserDisplayName] = () =>
                        result.UserDisplayName = principal.DisplayName,
                    [UserDistinguishedName] = () =>
                        result.UserDistinguishedName =
                            principal.DistinguishedName,
                    [UserDivision] = () =>
                        result.UserDivision = principal.GetDivision(),
                    [UserEmailAddress] = () =>
                        result.UserEmailAddress = principal.EmailAddress,
                    [UserEmployeeId] = () =>
                        result.UserEmployeeId = principal.EmployeeId,
                    [UserEnabled] = () =>
                        result.UserEnabled = principal.Enabled,
                    [UserFax] = () =>
                        result.UserFax = principal.GetFax(),
                    [UserSuffix] = () =>
                        result.UserSuffix = principal.GetSuffix(),
                    [UserGivenName] = () =>
                        result.UserGivenName = principal.GivenName,
                    [UserGuid] = () =>
                        result.UserGuid = principal.Guid,
                    [UserHomeAddress] = () =>
                        result.UserHomeAddress = principal.GetHomeAddress(),
                    [UserHomeDirectory] = () =>
                        result.UserHomeDirectory = principal.HomeDirectory,
                    [UserHomeDrive] = () =>
                        result.UserHomeDrive = principal.HomeDrive,
                    [UserHomePhone] = () =>
                        result.UserHomePhone = principal.GetHomePhone(),
                    [UserInitials] = () =>
                        result.UserInitials = principal.GetInitials(),
                    [UserIsAccountLockedOut] = () =>
                        result.UserIsAccountLockedOut =
                            principal.IsAccountLockedOut(),
                    [UserIsActive] = () =>
                        result.UserIsActive = principal.IsActive(),
                    [UserLastBadPasswordAttempt] = () =>
                        result.UserLastBadPasswordAttempt =
                            principal.LastBadPasswordAttempt,
                    [UserLastLogon] = () =>
                        result.UserLastLogon = principal.LastLogon,
                    [UserLastPasswordSet] = () =>
                        result.UserLastPasswordSet = principal.LastPasswordSet,
                    [UserManager] = () =>
                        result.UserManager = principal.GetManager(),
                    [UserMiddleName] = () =>
                        result.UserMiddleName = principal.MiddleName,
                    [UserMobile] = () =>
                        result.UserMobile = principal.GetMobile(),
                    [UserName] = () =>
                        result.UserName = principal.Name,
                    [UserNotes] = () =>
                        result.UserNotes = principal.GetNotes(),
                    [UserPager] = () =>
                        result.UserPager = principal.GetPager(),
                    [UserPasswordNeverExpires] = () =>
                        result.UserPasswordNeverExpires =
                            principal.PasswordNeverExpires,
                    [UserPasswordNotRequired] = () =>
                        result.UserPasswordNotRequired =
                            principal.PasswordNotRequired,
                    [UserPermittedLogonTimes] = () =>
                        result.UserPermittedLogonTimes =
                            principal.PermittedLogonTimes,
                    [UserPermittedWorkstations] = () =>
                        result.UserPermittedWorkstations =
                            principal.PermittedWorkstations,
                    [UserSamAccountName] = () =>
                        result.UserSamAccountName = principal.SamAccountName,
                    [UserScriptPath] = () =>
                        result.UserScriptPath = principal.ScriptPath,
                    [UserSid] = () =>
                        result.UserSid = principal.Sid,
                    [UserSip] = () =>
                        result.UserSip = principal.GetSip(),
                    [UserSmartcardLogonRequired] = () =>
                        result.UserSmartcardLogonRequired =
                            principal.SmartcardLogonRequired,
                    [UserState] = () =>
                        result.UserState = principal.GetState(),
                    [UserStreetAddress] = () =>
                        result.UserStreetAddress =
                            principal.GetStreetAddress(),
                    [UserStructuralObjectClass] = () =>
                        result.UserStructuralObjectClass =
                            principal.StructuralObjectClass,
                    [UserSurname] = () =>
                        result.UserSurname = principal.Surname,
                    [UserTitle] = () =>
                        result.UserTitle = principal.GetTitle(),
                    [UserUserCannotChangePassword] = () =>
                        result.UserUserCannotChangePassword =
                            principal.UserCannotChangePassword,
                    [UserUserPrincipalName] = () =>
                        result.UserUserPrincipalName =
                            principal.UserPrincipalName,
                    [UserVoiceTelephoneNumber] = () =>
                        result.UserVoiceTelephoneNumber =
                            principal.VoiceTelephoneNumber,
                    [UserVoip] = () =>
                        result.UserVoip = principal.GetVoip()
                };

            if (!attributeMapping.ContainsKey(attribute)) return;
            attributeMapping[attribute]();
        }

        private void AddAttributesToResult(
            dynamic result,
            ComputerPrincipal computerPrincipal = null,
            GroupPrincipal groupPrincipal = null,
            UserPrincipal directReportUserPrincipal = null,
            UserPrincipal userPrincipal = null)
        {
            foreach (var attribute in Attributes)
            {
                if (computerPrincipal != null)
                {
                    if (AddComputerAttributeToResult(
                        computerPrincipal, attribute, result))
                    {
                        continue;
                    }
                }
                if (directReportUserPrincipal != null)
                {
                    if (AddDirectReportAttributeToResult(
                        directReportUserPrincipal, attribute, result))
                    {
                        continue;
                    }
                }
                if (groupPrincipal != null)
                {
                    if (AddGroupAttributeToResult(
                        groupPrincipal, attribute, result))
                    {
                        continue;
                    }
                }
                if (userPrincipal != null)
                {
                    AddUserAttributeToResult(userPrincipal, attribute, result);
                }
            }
        }
    }
}