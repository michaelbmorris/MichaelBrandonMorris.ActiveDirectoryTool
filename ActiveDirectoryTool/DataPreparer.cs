using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Threading;
using Extensions.PrincipalExtensions;
using ActiveDirectoryPropertyMapping = System.Collections.Generic.Dictionary
    <ActiveDirectoryTool.ActiveDirectoryProperty, System.Action>;
using static ActiveDirectoryTool.ActiveDirectoryProperty;

namespace ActiveDirectoryTool
{
    public enum ActiveDirectoryProperty
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

    public class DataPreparer
    {
        public IEnumerable<ActiveDirectoryProperty> Properties { get; set; }
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
            foreach (var property in Properties)
            {
                if (AddComputerAttributeToResult(
                    computerPrincipal, property, result))
                {
                    continue;
                }

                if (AddContainerGroupAttributeToResult(
                    containerGroupPrincipal, property, result))
                {
                    continue;
                }

                if (AddDirectReportAttributeToResult(
                    directReportUserPrincipal, property, result))
                {
                    continue;
                }

                if (AddGroupAttributeToResult(
                    groupPrincipal, property, result))
                {
                    continue;
                }

                AddUserAttributeToResult(userPrincipal, property, result);
            }
        }

        private static bool AddComputerAttributeToResult(
            ComputerPrincipal computerPrincipal,
            ActiveDirectoryProperty property,
            dynamic result)
        {
            var propertyMapping = new ActiveDirectoryPropertyMapping
            {
                [ComputerAccountExpirationDate] = () =>
                {
                    result.ComputerAccountExpirationDate =
                        computerPrincipal?.AccountExpirationDate;
                },
                [ComputerAccountLockoutTime] = () =>
                {
                    result.ComputerAccountLockoutTime =
                        computerPrincipal?.AccountLockoutTime;
                },
                [ComputerAllowReversiblePasswordEncryption] = () =>
                {
                    result.ComputerAllowReversiblePasswordEncryption =
                        computerPrincipal?
                            .AllowReversiblePasswordEncryption;
                },
                [ComputerBadLogonCount] = () =>
                {
                    result.ComputerBadLogonCount =
                        computerPrincipal?.BadLogonCount;
                },
                [ComputerCertificates] = () =>
                {
                    result.ComputerCertificates =
                        computerPrincipal?.Certificates;
                },
                [ComputerContext] = () =>
                {
                    result.ComputerContext = computerPrincipal?.Context;
                },
                [ComputerContextType] = () =>
                {
                    result.ComputerContextType =
                        computerPrincipal?.ContextType;
                },
                [ComputerDelegationPermitted] = () =>
                {
                    result.ComputerDelegationPermitted =
                        computerPrincipal?.DelegationPermitted;
                },
                [ComputerDescription] = () =>
                {
                    result.ComputerDescription =
                        computerPrincipal?.Description;
                },
                [ComputerDisplayName] = () =>
                {
                    result.ComputerDisplayName =
                        computerPrincipal?.DisplayName;
                },
                [ComputerDistinguishedName] = () =>
                {
                    result.ComputerDistinguishedName =
                        computerPrincipal?.DistinguishedName;
                },
                [ComputerEnabled] = () =>
                {
                    result.ComputerEnabled = computerPrincipal?.Enabled;
                },
                [ComputerGuid] = () =>
                {
                    result.ComputerGuid = computerPrincipal?.Guid;
                },
                [ComputerHomeDirectory] = () =>
                {
                    result.ComputerHomeDirectory =
                        computerPrincipal?.HomeDirectory;
                },
                [ComputerHomeDrive] = () =>
                {
                    result.ComputerHomeDrive =
                        computerPrincipal?.HomeDrive;
                },
                [ComputerLastBadPasswordAttempt] = () =>
                {
                    result.ComputerLastBadPasswordAttempt =
                        computerPrincipal?.LastBadPasswordAttempt;
                },
                [ComputerLastLogon] = () =>
                {
                    result.ComputerLastLogon =
                        computerPrincipal?.LastLogon;
                },
                [ComputerLastPasswordSet] = () =>
                {
                    result.LastPasswordSet =
                        computerPrincipal?.LastPasswordSet;
                },
                [ComputerName] = () =>
                {
                    result.ComputerName = computerPrincipal?.Name;
                },
                [ComputerPasswordNeverExpires] = () =>
                {
                    result.ComputerPasswordNeverExpires =
                        computerPrincipal?.PasswordNeverExpires;
                },
                [ComputerPasswordNotRequired] = () =>
                {
                    result.ComputerPasswordNotRequired =
                        computerPrincipal?.PasswordNotRequired;
                },
                [ComputerPermittedLogonTimes] = () =>
                {
                    result.ComputerPermittedLogonTimes =
                        computerPrincipal?.PermittedLogonTimes;
                },
                [ComputerPermittedWorkstations] = () =>
                {
                    result.ComputerPermittedWorkstations =
                        computerPrincipal?.PermittedWorkstations;
                },
                [ComputerSamAccountName] = () =>
                {
                    result.ComputerSamAccountName =
                        computerPrincipal?.SamAccountName;
                },
                [ComputerScriptPath] = () =>
                {
                    result.ComputerScriptPath =
                        computerPrincipal?.ScriptPath;
                },
                [ComputerServicePrincipalNames] = () =>
                {
                    result.ComputerServicecomputerPrincipalNames =
                        computerPrincipal?.ServicePrincipalNames;
                },
                [ComputerSid] = () =>
                {
                    result.ComputerSid = computerPrincipal?.Sid;
                },
                [ComputerSmartcardLogonRequired] = () =>
                {
                    result.ComputerSmartcardLogonRequired =
                        computerPrincipal?.SmartcardLogonRequired;
                },
                [ComputerStructuralObjectClass] = () =>
                {
                    result.ComputerStructuralObjectClass =
                        computerPrincipal?.StructuralObjectClass;
                },
                [ComputerUserCannotChangePassword] = () =>
                {
                    result.ComputerUserCannotChangePassword =
                        computerPrincipal?.UserCannotChangePassword;
                },
                [ComputerUserPrincipalName] = () =>
                {
                    result.ComputerUsercomputerPrincipalName =
                        computerPrincipal?.UserPrincipalName;
                }
            };

            if (!propertyMapping.ContainsKey(property)) return false;
            propertyMapping[property]();
            return true;
        }

        private static bool AddContainerGroupAttributeToResult(
            GroupPrincipal containerGroupPrincipal,
            ActiveDirectoryProperty property,
            dynamic result)
        {
            var propertyMapping = new ActiveDirectoryPropertyMapping
            {
                [ContainerGroupContext] = () =>
                {
                    result.ContainerGroupContext = 
                        containerGroupPrincipal?.Context;
                },
                [ContainerGroupContextType] = () =>
                {
                    result.ContainerGroupContextType =
                        containerGroupPrincipal?.ContextType;
                },
                [ContainerGroupDescription] = () =>
                {
                    result.ContainerGroupDescription =
                        containerGroupPrincipal?.Description;
                },
                [ContainerGroupDisplayName] = () =>
                {
                    result.ContainerGroupDisplayName =
                        containerGroupPrincipal?.DisplayName;
                },
                [ContainerGroupDistinguishedName] = () =>
                {
                    result.ContainerGroupDistinguishedName =
                        containerGroupPrincipal?.DistinguishedName;
                },
                [ContainerGroupGuid] = () =>
                {
                    result.ContainerGroupGuid = containerGroupPrincipal?.Guid;
                },
                [ContainerGroupIsSecurityGroup] = () =>
                {
                    result.ContainerGroupIsSecurityGroup =
                        containerGroupPrincipal?.IsSecurityGroup;
                },
                [ContainerGroupManagedBy] = () =>
                {
                    result.ContainerGroupManagedBy =
                        containerGroupPrincipal?.GetManagedBy();
                },
                [ContainerGroupName] = () =>
                {
                    result.ContainerGroupName = containerGroupPrincipal?.Name;
                },
                [ContainerGroupSamAccountName] = () =>
                {
                    result.ContainerGroupSamAccountName =
                        containerGroupPrincipal?.SamAccountName;
                },
                [ContainerGroupScope] = () =>
                {
                    result.ContainerGroupScope = 
                        containerGroupPrincipal?.GroupScope;
                },
                [ContainerGroupSid] = () =>
                {
                    result.ContainerGroupSid = containerGroupPrincipal?.Sid;
                },
                [ContainerGroupStructuralObjectClass] = () =>
                {
                    result.ContainerGroupStructuralObjectClass =
                        containerGroupPrincipal?.StructuralObjectClass;
                },
                [ContainerGroupUserPrincipalName] = () =>
                {
                    result.ContainerGroupUserPrincipalName =
                        containerGroupPrincipal?.UserPrincipalName;
                },
                [ContainerGroupMembers] = () =>
                {
                    result.ContainerGroupMembers = 
                        containerGroupPrincipal?.Members;
                }
            };

            if (!propertyMapping.ContainsKey(property)) return false;
            propertyMapping[property]();
            return true;
        }

        private static bool AddDirectReportAttributeToResult(
            UserPrincipal directReportUserPrincipal,
            ActiveDirectoryProperty property,
            dynamic result)
        {
            var propertyMapping = new ActiveDirectoryPropertyMapping
            {
                [DirectReportUserAccountControl] = () =>
                    result.DirectReportAccountControl =
                        directReportUserPrincipal.GetUserAccountControl(),
                [DirectReportAccountExpirationDate] = () =>
                    result.DirectReportAccountExpirationDate =
                        directReportUserPrincipal.AccountExpirationDate,
                [DirectReportAccountLockoutTime] = () =>
                    result.DirectReportAccountLockoutTime =
                        directReportUserPrincipal.AccountLockoutTime,
                [DirectReportAllowReversiblePasswordEncryption] = () =>
                    result.DirectReportAllowReversiblePasswordEncryption =
                        directReportUserPrincipal
                            .AllowReversiblePasswordEncryption,
                [DirectReportAssistant] = () =>
                    result.DirectReportAssistant =
                        directReportUserPrincipal.GetAssistant(),
                [DirectReportBadLogonCount] = () =>
                    result.DirectReportBadLogonCount =
                        directReportUserPrincipal.BadLogonCount,
                [DirectReportCertificates] = () =>
                    result.DirectReportCertificates =
                        directReportUserPrincipal.Certificates,
                [DirectReportCity] = () =>
                    result.DirectReportCity =
                        directReportUserPrincipal.GetCity(),
                [DirectReportComment] = () =>
                    result.DirectReportComment =
                        directReportUserPrincipal.GetComment(),
                [DirectReportCompany] = () =>
                    result.DirectReportCompany =
                        directReportUserPrincipal.GetCompany(),
                [DirectReportContext] = () =>
                    result.DirectReportContext =
                        directReportUserPrincipal.Context,
                [DirectReportContextType] = () =>
                    result.DirectReportContextType =
                        directReportUserPrincipal.ContextType,
                [DirectReportCountry] = () =>
                    result.DirectReportCountry =
                        directReportUserPrincipal.GetCountry(),
                [DirectReportDelegationPermitted] = () =>
                    result.DirectReportDelegationPermitted =
                        directReportUserPrincipal.DelegationPermitted,
                [DirectReportDepartment] = () =>
                    result.DirectReportDepartment =
                        directReportUserPrincipal.GetDepartment(),
                [DirectReportDescription] = () =>
                    result.DirectReportDescription =
                        directReportUserPrincipal.Description,
                [DirectReportDisplayName] = () =>
                    result.DirectReportDisplayName =
                        directReportUserPrincipal.DisplayName,
                [DirectReportDistinguishedName] = () =>
                {
                    result.DirectReportDistinguishedName =
                        directReportUserPrincipal?.DistinguishedName;
                },
                [DirectReportDivision] = () =>
                    result.DirectReportDivision =
                        directReportUserPrincipal.GetDivision(),
                [DirectReportEmailAddress] = () =>
                    result.DirectReportEmailAddress =
                        directReportUserPrincipal.EmailAddress,
                [DirectReportEmployeeId] = () =>
                    result.DirectReportEmployeeId =
                        directReportUserPrincipal.EmployeeId,
                [DirectReportEnabled] = () =>
                    result.DirectReportEnabled =
                        directReportUserPrincipal.Enabled,
                [DirectReportFax] = () =>
                    result.DirectReportFax =
                        directReportUserPrincipal.GetFax(),
                [DirectReportSuffix] = () =>
                    result.DirectReportSuffix =
                        directReportUserPrincipal.GetSuffix(),
                [DirectReportGivenName] = () =>
                    result.DirectReportGivenName =
                        directReportUserPrincipal.GivenName,
                [DirectReportGuid] = () =>
                    result.DirectReportGuid =
                        directReportUserPrincipal.Guid,
                [DirectReportHomeAddress] = () =>
                    result.DirectReportHomeAddress =
                        directReportUserPrincipal.GetHomeAddress(),
                [DirectReportHomeDirectory] = () =>
                    result.DirectReportHomeDirectory =
                        directReportUserPrincipal.HomeDirectory,
                [DirectReportHomeDrive] = () =>
                    result.DirectReportHomeDrive =
                        directReportUserPrincipal.HomeDrive,
                [DirectReportHomePhone] = () =>
                    result.DirectReportHomePhone =
                        directReportUserPrincipal.GetHomePhone(),
                [DirectReportInitials] = () =>
                    result.DirectReportInitials =
                        directReportUserPrincipal.GetInitials(),
                [DirectReportIsAccountLockedOut] = () =>
                    result.DirectReportIsAccountLockedOut =
                        directReportUserPrincipal.IsAccountLockedOut(),
                [DirectReportIsActive] = () =>
                    result.DirectReportIsActive =
                        directReportUserPrincipal.IsActive(),
                [DirectReportLastBadPasswordAttempt] = () =>
                    result.DirectReportLastBadPasswordAttempt =
                        directReportUserPrincipal.LastBadPasswordAttempt,
                [DirectReportLastLogon] = () =>
                    result.DirectReportLastLogon =
                        directReportUserPrincipal.LastLogon,
                [DirectReportLastPasswordSet] = () =>
                    result.DirectReportLastPasswordSet =
                        directReportUserPrincipal.LastPasswordSet,
                [DirectReportManager] = () =>
                    result.DirectReportManager =
                        directReportUserPrincipal.GetManager(),
                [DirectReportMiddleName] = () =>
                    result.DirectReportMiddleName =
                        directReportUserPrincipal.MiddleName,
                [DirectReportMobile] = () =>
                    result.DirectReportMobile =
                        directReportUserPrincipal.GetMobile(),
                [DirectReportName] = () =>
                {
                    result.DirectReportName =
                        directReportUserPrincipal?.Name;
                },
                [DirectReportNotes] = () =>
                    result.DirectReportNotes =
                        directReportUserPrincipal.GetNotes(),
                [DirectReportPager] = () =>
                    result.DirectReportPager =
                        directReportUserPrincipal.GetPager(),
                [DirectReportPasswordNeverExpires] = () =>
                    result.DirectReportPasswordNeverExpires =
                        directReportUserPrincipal.PasswordNeverExpires,
                [DirectReportPasswordNotRequired] = () =>
                    result.DirectReportPasswordNotRequired =
                        directReportUserPrincipal.PasswordNotRequired,
                [DirectReportPermittedLogonTimes] = () =>
                    result.DirectReportPermittedLogonTimes =
                        directReportUserPrincipal.PermittedLogonTimes,
                [DirectReportPermittedWorkstations] = () =>
                    result.DirectReportPermittedWorkstations =
                        directReportUserPrincipal.PermittedWorkstations,
                [DirectReportSamAccountName] = () =>
                    result.DirectReportSamAccountName =
                        directReportUserPrincipal.SamAccountName,
                [DirectReportScriptPath] = () =>
                    result.DirectReportScriptPath =
                        directReportUserPrincipal.ScriptPath,
                [DirectReportSid] = () =>
                    result.DirectReportSid = directReportUserPrincipal.Sid,
                [DirectReportSip] = () =>
                    result.DirectReportSip =
                        directReportUserPrincipal.GetSip(),
                [DirectReportSmartcardLogonRequired] = () =>
                    result.DirectReportSmartcardLogonRequired =
                        directReportUserPrincipal.SmartcardLogonRequired,
                [DirectReportState] = () =>
                    result.DirectReportState =
                        directReportUserPrincipal.GetState(),
                [DirectReportStreetAddress] = () =>
                    result.DirectReportStreetAddress =
                        directReportUserPrincipal.GetStreetAddress(),
                [DirectReportStructuralObjectClass] = () =>
                    result.DirectReportStructuralObjectClass =
                        directReportUserPrincipal.StructuralObjectClass,
                [DirectReportSurname] = () =>
                    result.DirectReportSurname =
                        directReportUserPrincipal.Surname,
                [DirectReportTitle] = () =>
                    result.DirectReportTitle =
                        directReportUserPrincipal.GetTitle(),
                [DirectReportUserCannotChangePassword] = () =>
                    result.DirectReportUserCannotChangePassword =
                        directReportUserPrincipal.UserCannotChangePassword,
                [DirectReportUserPrincipalName] = () =>
                    result.DirectReportdirectReportUserPrincipalName =
                        directReportUserPrincipal.UserPrincipalName,
                [DirectReportVoiceTelephoneNumber] = () =>
                    result.DirectReportVoiceTelephoneNumber =
                        directReportUserPrincipal.VoiceTelephoneNumber,
                [DirectReportVoip] = () =>
                    result.DirectReportVoip =
                        directReportUserPrincipal.GetVoip()
            };

            if (!propertyMapping.ContainsKey(property)) return false;
            propertyMapping[property]();
            return true;
        }

        private static bool AddGroupAttributeToResult(
            GroupPrincipal principal,
            ActiveDirectoryProperty property,
            dynamic result)
        {
            var propertyMapping = new ActiveDirectoryPropertyMapping
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
                [ActiveDirectoryProperty.GroupScope] = () =>
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

            if (!propertyMapping.ContainsKey(property)) return false;
            propertyMapping[property]();
            return true;
        }

        private static void AddUserAttributeToResult(
            UserPrincipal principal,
            ActiveDirectoryProperty property,
            dynamic result)
        {
            var propertyMapping = new ActiveDirectoryPropertyMapping
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

            if (!propertyMapping.ContainsKey(property)) return;
            propertyMapping[property]();
        }

        public IEnumerable<ExpandoObject> GetResults()
        {
            var results = new List<ExpandoObject>();

            foreach (var data in Data.Value)
            {
                CancellationToken.ThrowIfCancellationRequested();
                if (data is User)
                {
                    var user = data as User;
                    dynamic result = new ExpandoObject();
                    result.Name = user.Name;
                    results.Add(result);
                }
                else if (data is ComputerGroups)
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