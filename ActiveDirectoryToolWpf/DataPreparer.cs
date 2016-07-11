using System;
using System.Collections.Generic;
using System.Data;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Linq;
using System.Threading;
using CollectionExtensions;
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
        ContainerGroupContext,
        ContainerGroupContextType,
        ContainerGroupDescription,
        ContainerGroupDisplayName,
        ContainerGroupDistinguishedName,
        ContainerGroupGuid,
        ContainerGroupIsSecurityGroup,
        ContainerGroupManagedBy,
        ContainerGroupMembers,
        ContainerGroupName,
        ContainerGroupSamAccountName,
        ContainerGroupScope,
        ContainerGroupSid,
        ContainerGroupStructuralObjectClass,
        ContainerGroupUserPrincipalName,
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
        public static DataTable ToDataTable(
            this IEnumerable<ExpandoObject> data)
        {
            var dataAsArray = data as ExpandoObject[] ?? data.ToArray();
            if (dataAsArray.IsNullOrEmpty()) return null;
            var dataTable = new DataTable();
            foreach (var dataItem in dataAsArray)
            {
                foreach (var key in ((IDictionary<string, object>) dataItem)
                    .Keys)
                {
                    if(!dataTable.Columns.Contains(key))
                        dataTable.Columns.Add(key);
                }
            }
            
            foreach (var dataItem in dataAsArray)
            {
                dataTable.Rows.Add(
                    ((IDictionary<string, object>) dataItem)
                        .Values.ToArray());
            }
            return dataTable;
        }
    }

    public class DataPreparer
    {
        public IEnumerable<ActiveDirectoryAttribute> Attributes { get; set; }

        public CancellationToken CancellationToken { get; set; }

        public Lazy<IEnumerable<object>> Data { get; set; }

        private void AddAttributesToResult(
            dynamic result,
            ComputerPrincipal computerPrincipal = null,
            GroupPrincipal groupPrincipal = null,
            GroupPrincipal containerGroupPrincipal = null,
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
                if (containerGroupPrincipal != null)
                {
                    if (AddContainerGroupAttributeToResult(
                        containerGroupPrincipal, attribute, result))
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

        private static bool AddContainerGroupAttributeToResult(
            GroupPrincipal principal,
            ActiveDirectoryAttribute attribute,
            dynamic result)
        {
            var attributeMapping =
                new Dictionary<ActiveDirectoryAttribute, Action>
                {
                    [ContainerGroupContext] =
                        () => result.ContainerGroupContext = principal.Context,
                    [ContainerGroupContextType] = () =>
                        result.ContainerGroupContextType =
                            principal.ContextType,
                    [ContainerGroupDescription] = () =>
                        result.ContainerGroupDescription =
                            principal.Description,
                    [ContainerGroupDisplayName] = () =>
                        result.ContainerGroupDisplayName =
                            principal.DisplayName,
                    [ContainerGroupDistinguishedName] = () =>
                        result.ContainerGroupDistinguishedName =
                            principal.DistinguishedName,
                    [ContainerGroupGuid] = () => result.ContainerGroupGuid =
                        principal.Guid,
                    [ContainerGroupIsSecurityGroup] = () =>
                        result.ContainerGroupIsSecurityGroup =
                            principal.IsSecurityGroup,
                    [ContainerGroupManagedBy] = () =>
                        result.ContainerGroupManagedBy =
                            principal.GetManagedBy(),
                    [ContainerGroupName] =
                        () => result.ContainerGroupName = principal.Name,
                    [ContainerGroupSamAccountName] = () =>
                        result.ContainerGroupSamAccountName =
                            principal.SamAccountName,
                    [ContainerGroupScope] = () =>
                        result.ContainerGroupScope = principal.GroupScope,
                    [ContainerGroupSid] =
                        () => result.ContainerGroupSid = principal.Sid,
                    [ContainerGroupStructuralObjectClass] = () =>
                        result.ContainerGroupStructuralObjectClass =
                            principal.StructuralObjectClass,
                    [ContainerGroupUserPrincipalName] = () =>
                        result.ContainerGroupUserPrincipalName =
                            principal.UserPrincipalName,
                    [ContainerGroupMembers] =
                        () => result.ContainerGroupMembers = principal.Members
                };

            if (!attributeMapping.ContainsKey(attribute)) return false;
            attributeMapping[attribute]();
            return true;
        }

        private static bool AddDirectReportAttributeToResult(
            UserPrincipal userPrincipal,
            ActiveDirectoryAttribute attribute,
            dynamic result)
        {
            var attributeMapping =
                new Dictionary<ActiveDirectoryAttribute, Action>
                {
                    [DirectReportUserAccountControl] = () =>
                        result.DirectReportAccountControl =
                            userPrincipal.GetUserAccountControl(),
                    [DirectReportAccountExpirationDate] = () =>
                        result.DirectReportAccountExpirationDate =
                            userPrincipal.AccountExpirationDate,
                    [DirectReportAccountLockoutTime] = () =>
                        result.DirectReportAccountLockoutTime =
                            userPrincipal.AccountLockoutTime,
                    [DirectReportAllowReversiblePasswordEncryption] = () =>
                        result.DirectReportAllowReversiblePasswordEncryption =
                            userPrincipal.AllowReversiblePasswordEncryption,
                    [DirectReportAssistant] = () =>
                        result.DirectReportAssistant =
                            userPrincipal.GetAssistant(),
                    [DirectReportBadLogonCount] = () =>
                        result.DirectReportBadLogonCount =
                            userPrincipal.BadLogonCount,
                    [DirectReportCertificates] = () =>
                        result.DirectReportCertificates =
                            userPrincipal.Certificates,
                    [DirectReportCity] = () =>
                        result.DirectReportCity = userPrincipal.GetCity(),
                    [DirectReportComment] = () =>
                        result.DirectReportComment = userPrincipal.GetComment(),
                    [DirectReportCompany] = () =>
                        result.DirectReportCompany = userPrincipal.GetCompany(),
                    [DirectReportContext] = () =>
                        result.DirectReportContext = userPrincipal.Context,
                    [DirectReportContextType] = () =>
                        result.DirectReportContextType = userPrincipal.ContextType,
                    [DirectReportCountry] = () =>
                        result.DirectReportCountry = userPrincipal.GetCountry(),
                    [DirectReportDelegationPermitted] = () =>
                        result.DirectReportDelegationPermitted =
                            userPrincipal.DelegationPermitted,
                    [DirectReportDepartment] = () =>
                        result.DirectReportDepartment =
                            userPrincipal.GetDepartment(),
                    [DirectReportDescription] = () =>
                        result.DirectReportDescription = userPrincipal.Description,
                    [DirectReportDisplayName] = () =>
                        result.DirectReportDisplayName = userPrincipal.DisplayName,
                    [DirectReportDistinguishedName] = () =>
                        result.DirectReportDistinguishedName =
                            userPrincipal.DistinguishedName,
                    [DirectReportDivision] = () =>
                        result.DirectReportDivision = userPrincipal.GetDivision(),
                    [DirectReportEmailAddress] = () =>
                        result.DirectReportEmailAddress =
                            userPrincipal.EmailAddress,
                    [DirectReportEmployeeId] = () =>
                        result.DirectReportEmployeeId = userPrincipal.EmployeeId,
                    [DirectReportEnabled] = () =>
                        result.DirectReportEnabled = userPrincipal.Enabled,
                    [DirectReportFax] = () =>
                        result.DirectReportFax = userPrincipal.GetFax(),
                    [DirectReportSuffix] = () =>
                        result.DirectReportSuffix = userPrincipal.GetSuffix(),
                    [DirectReportGivenName] = () =>
                        result.DirectReportGivenName = userPrincipal.GivenName,
                    [DirectReportGuid] = () =>
                        result.DirectReportGuid = userPrincipal.Guid,
                    [DirectReportHomeAddress] = () =>
                        result.DirectReportHomeAddress =
                            userPrincipal.GetHomeAddress(),
                    [DirectReportHomeDirectory] = () =>
                        result.DirectReportHomeDirectory =
                            userPrincipal.HomeDirectory,
                    [DirectReportHomeDrive] = () =>
                        result.DirectReportHomeDrive = userPrincipal.HomeDrive,
                    [DirectReportHomePhone] = () =>
                        result.DirectReportHomePhone =
                            userPrincipal.GetHomePhone(),
                    [DirectReportInitials] = () =>
                        result.DirectReportInitials = userPrincipal.GetInitials(),
                    [DirectReportIsAccountLockedOut] = () =>
                        result.DirectReportIsAccountLockedOut =
                            userPrincipal.IsAccountLockedOut(),
                    [DirectReportIsActive] = () =>
                        result.DirectReportIsActive = userPrincipal.IsActive(),
                    [DirectReportLastBadPasswordAttempt] = () =>
                        result.DirectReportLastBadPasswordAttempt =
                            userPrincipal.LastBadPasswordAttempt,
                    [DirectReportLastLogon] = () =>
                        result.DirectReportLastLogon = userPrincipal.LastLogon,
                    [DirectReportLastPasswordSet] = () =>
                        result.DirectReportLastPasswordSet =
                            userPrincipal.LastPasswordSet,
                    [DirectReportManager] = () =>
                        result.DirectReportManager = userPrincipal.GetManager(),
                    [DirectReportMiddleName] = () =>
                        result.DirectReportMiddleName = userPrincipal.MiddleName,
                    [DirectReportMobile] = () =>
                        result.DirectReportMobile = userPrincipal.GetMobile(),
                    [DirectReportName] = () => result.DirectReportName =
                        userPrincipal.Name,
                    [DirectReportNotes] = () =>
                        result.DirectReportNotes = userPrincipal.GetNotes(),
                    [DirectReportPager] = () =>
                        result.DirectReportPager = userPrincipal.GetPager(),
                    [DirectReportPasswordNeverExpires] = () =>
                        result.DirectReportPasswordNeverExpires =
                            userPrincipal.PasswordNeverExpires,
                    [DirectReportPasswordNotRequired] = () =>
                        result.DirectReportPasswordNotRequired =
                            userPrincipal.PasswordNotRequired,
                    [DirectReportPermittedLogonTimes] = () =>
                        result.DirectReportPermittedLogonTimes =
                            userPrincipal.PermittedLogonTimes,
                    [DirectReportPermittedWorkstations] = () =>
                        result.DirectReportPermittedWorkstations =
                            userPrincipal.PermittedWorkstations,
                    [DirectReportSamAccountName] = () =>
                        result.DirectReportSamAccountName =
                            userPrincipal.SamAccountName,
                    [DirectReportScriptPath] = () =>
                        result.DirectReportScriptPath = userPrincipal.ScriptPath,
                    [DirectReportSid] = () =>
                        result.DirectReportSid = userPrincipal.Sid,
                    [DirectReportSip] = () =>
                        result.DirectReportSip = userPrincipal.GetSip(),
                    [DirectReportSmartcardLogonRequired] = () =>
                        result.DirectReportSmartcardLogonRequired =
                            userPrincipal.SmartcardLogonRequired,
                    [DirectReportState] = () =>
                        result.DirectReportState = userPrincipal.GetState(),
                    [DirectReportStreetAddress] = () =>
                        result.DirectReportStreetAddress =
                            userPrincipal.GetStreetAddress(),
                    [DirectReportStructuralObjectClass] = () =>
                        result.DirectReportStructuralObjectClass =
                            userPrincipal.StructuralObjectClass,
                    [DirectReportSurname] = () =>
                        result.DirectReportSurname = userPrincipal.Surname,
                    [DirectReportTitle] = () =>
                        result.DirectReportTitle = userPrincipal.GetTitle(),
                    [DirectReportUserCannotChangePassword] = () =>
                        result.DirectReportUserCannotChangePassword =
                            userPrincipal.UserCannotChangePassword,
                    [DirectReportUserPrincipalName] = () =>
                        result.DirectReportUserPrincipalName =
                            userPrincipal.UserPrincipalName,
                    [DirectReportVoiceTelephoneNumber] = () =>
                        result.DirectReportVoiceTelephoneNumber =
                            userPrincipal.VoiceTelephoneNumber,
                    [DirectReportVoip] = () =>
                        result.DirectReportVoip = userPrincipal.GetVoip()
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

        public IEnumerable<ExpandoObject> GetResults()
        {
            var results = new List<ExpandoObject>();

            foreach (var data in Data.Value)
            {
                CancellationToken.ThrowIfCancellationRequested();
                if (data is ComputerGroups)
                {
                    var computerGroups = data as ComputerGroups;
                    results.AddRange(PrepareComputerGroups(computerGroups));
                }
                else if (data is GroupComputers)
                {
                    var groupComputers = data as GroupComputers;
                    results.AddRange(PrepareGroupComputers(groupComputers));
                }
                else if (data is GroupUsers)
                {
                    var groupUsers = data as GroupUsers;
                    results.AddRange(PrepareGroupUsers(groupUsers));
                }
                else if (data is GroupUsersDirectReports)
                {
                    var groupUsersDirectReports =
                        data as GroupUsersDirectReports;
                    results.AddRange(
                        PrepareGroupUsersDirectReports(
                            groupUsersDirectReports));
                }
                else if (data is GroupUsersGroups)
                {
                    var groupUsersGroups = data as GroupUsersGroups;
                    results.AddRange(
                        PrepareGroupUsersGroups(groupUsersGroups));
                }
                else if (data is UserGroups)
                {
                    var userGroups = data as UserGroups;
                    results.AddRange(PrepareUserGroups(userGroups));
                }
                else if (data is UserDirectReports)
                {
                    var userDirectReports = data as UserDirectReports;
                    results.AddRange(
                        PrepareUserDirectReports(
                            userDirectReports));
                }
                else
                {
                    var computerPrincipal = data as ComputerPrincipal;
                    var groupPrincipal = data as GroupPrincipal;
                    var userPrincipal = data as UserPrincipal;
                    dynamic result = new ExpandoObject();
                    AddAttributesToResult(
                        result,
                        computerPrincipal: computerPrincipal,
                        groupPrincipal: groupPrincipal,
                        userPrincipal: userPrincipal);
                    results.Add(result);
                    var principal = data as Principal;
                    principal?.Dispose();
                }
            }
            return results;
        }

        private IEnumerable<ExpandoObject> PrepareComputerGroups(
            ComputerGroups computerGroups)
        {
            var results = new List<ExpandoObject>();
            if (computerGroups.Groups == null) return results;
            foreach (var group in computerGroups.Groups)
            {
                CancellationToken.ThrowIfCancellationRequested();
                dynamic result = new ExpandoObject();
                AddAttributesToResult(
                    result,
                    computerPrincipal: computerGroups.Computer,
                    groupPrincipal: group);
                results.Add(result);
            }
            return results;
        }

        private IEnumerable<ExpandoObject> PrepareGroupComputers(
            GroupComputers groupComputers)
        {
            var results = new List<ExpandoObject>();
            if (groupComputers.Computers == null) return results;
            foreach (var computer in groupComputers.Computers)
            {
                CancellationToken.ThrowIfCancellationRequested();
                dynamic result = new ExpandoObject();
                AddAttributesToResult(
                    result,
                    computerPrincipal: computer,
                    containerGroupPrincipal: groupComputers.Group);
                results.Add(result);
            }
            return results;
        }

        private IEnumerable<ExpandoObject> PrepareGroupUsers(
            GroupUsers groupUsers)
        {
            var results = new List<ExpandoObject>();
            if (groupUsers.Users == null) return results;
            foreach (var user in groupUsers.Users)
            {
                CancellationToken.ThrowIfCancellationRequested();
                dynamic result = new ExpandoObject();
                AddAttributesToResult(
                    result,
                    userPrincipal: user,
                    containerGroupPrincipal: groupUsers.Group);
                results.Add(result);
            }
            return results;
        }

        private IEnumerable<ExpandoObject> PrepareGroupUsersDirectReports(
            GroupUsersDirectReports groupUsersDirectReports)
        {
            var results = new List<ExpandoObject>();
            if (groupUsersDirectReports.UsersDirectReports == null)
                return results;
            foreach (var userDirectReports in groupUsersDirectReports
                .UsersDirectReports)
            {
                CancellationToken.ThrowIfCancellationRequested();
                if (userDirectReports.DirectReports == null) continue;
                foreach (var directReport in userDirectReports.DirectReports)
                {
                    CancellationToken.ThrowIfCancellationRequested();
                    dynamic result = new ExpandoObject();
                    AddAttributesToResult(
                        result,
                        containerGroupPrincipal: groupUsersDirectReports.Group,
                        userPrincipal: userDirectReports.User,
                        directReportUserPrincipal: directReport);
                    results.Add(result);
                }
            }
            return results;
        }

        private IEnumerable<ExpandoObject> PrepareGroupUsersGroups(
            GroupUsersGroups groupUsersGroups)
        {
            var results = new List<ExpandoObject>();
            if (groupUsersGroups.UsersGroups == null) return results;
            foreach (var userGroups in groupUsersGroups.UsersGroups)
            {
                CancellationToken.ThrowIfCancellationRequested();
                if (userGroups.Groups == null) continue;
                foreach (var group in userGroups.Groups)
                {
                    CancellationToken.ThrowIfCancellationRequested();
                    dynamic result = new ExpandoObject();
                    AddAttributesToResult(
                        result,
                        containerGroupPrincipal: groupUsersGroups.Group,
                        userPrincipal: userGroups.User,
                        groupPrincipal: group);
                    results.Add(result);
                }
            }
            return results;
        }

        private IEnumerable<ExpandoObject> PrepareUserDirectReports(
            UserDirectReports userDirectReports)
        {
            var results = new List<ExpandoObject>();
            if (userDirectReports.DirectReports == null) return results;
            foreach (var directReport in userDirectReports.DirectReports)
            {
                CancellationToken.ThrowIfCancellationRequested();
                if (directReport == null) continue;
                dynamic result = new ExpandoObject();
                AddAttributesToResult(
                    result,
                    directReportUserPrincipal: directReport,
                    userPrincipal: userDirectReports.User);
                results.Add(result);
            }
            return results;
        }

        private IEnumerable<ExpandoObject> PrepareUserGroups(
            UserGroups userGroups)
        {
            var results = new List<ExpandoObject>();
            if (userGroups.Groups == null) return results;
            foreach (var group in userGroups.Groups)
            {
                CancellationToken.ThrowIfCancellationRequested();
                dynamic result = new ExpandoObject();
                AddAttributesToResult(
                    result,
                    groupPrincipal: group,
                    userPrincipal: userGroups.User);
                results.Add(result);
            }
            return results;
        }
    }
}