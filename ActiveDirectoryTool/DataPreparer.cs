using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Linq;
using System.Threading;
using Extensions.PrincipalExtensions;
using ActiveDirectoryPropertyMapping = System.Collections.Generic.Dictionary
    <ActiveDirectoryTool.ActiveDirectoryProperty, System.Action>;
using static ActiveDirectoryTool.ActiveDirectoryProperty;

namespace ActiveDirectoryTool
{
    public class DataPreparer
    {
        private static readonly ActiveDirectoryProperty[]
            DefaultComputerGroupsProperties =
            {
                ComputerName,
                GroupName,
                ComputerDistinguishedName,
                GroupDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultComputerProperties =
            {
                ComputerName,
                ComputerDescription,
                ComputerDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultGroupComputersProperties =
            {
                GroupName,
                ComputerName,
                GroupDistinguishedName,
                ComputerDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultGroupProperties =
            {
                GroupName,
                GroupManagedBy,
                GroupDescription,
                GroupDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultGroupUsersDirectReportsProperties =
            {
                ContainerGroupName,
                UserName,
                DirectReportName,
                ContainerGroupDistinguishedName,
                UserDistinguishedName,
                DirectReportDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultGroupUsersGroupsProperties =
            {
                ContainerGroupName,
                UserName,
                GroupName,
                ContainerGroupDistinguishedName,
                UserDistinguishedName,
                GroupDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultGroupUsersProperties =
            {
                ContainerGroupName,
                UserName,
                ContainerGroupDistinguishedName,
                UserDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultUserDirectReportsProperties =
            {
                UserName,
                DirectReportName,
                UserDistinguishedName,
                DirectReportDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultUserGroupsProperties =
            {
                UserName,
                GroupName,
                UserDistinguishedName,
                GroupDistinguishedName
            };

        private static readonly ActiveDirectoryProperty[]
            DefaultUserProperties =
            {
                UserSurname,
                UserGivenName,
                UserDisplayName,
                UserSamAccountName,
                UserIsActive,
                UserIsAccountLockedOut,
                UserLastLogon,
                UserDescription,
                UserTitle,
                UserCompany,
                UserManager,
                UserHomeDrive,
                UserHomeDirectory,
                UserScriptPath,
                UserEmailAddress,
                UserStreetAddress,
                UserCity,
                UserState,
                UserVoiceTelephoneNumber,
                UserPager,
                UserMobile,
                UserFax,
                UserVoip,
                UserSip,
                UserUserPrincipalName,
                UserDistinguishedName
            };

        private readonly CancellationToken _cancellationToken;
        private readonly IEnumerable<string> _distinguishedNames;
        private readonly QueryType _queryType;
        private readonly Scope _scope;

        private IEnumerable<ActiveDirectoryProperty> _properties;

        public DataPreparer(
            QueryType queryType,
            Scope scope,
            IEnumerable<string> distinguishedNames,
            CancellationToken cancellationToken)
        {
            _queryType = queryType;
            _scope = scope;
            _distinguishedNames = distinguishedNames;
            _cancellationToken = cancellationToken;
        }

        public IEnumerable<ExpandoObject> GetData()
        {
            // ReSharper disable AccessToDisposedClosure
            IEnumerable<ExpandoObject> data;
            using (var principalContext = new PrincipalContext(
                ContextType.Domain,
                _scope.Domain,
                _scope.Context))
            {
                var mapping = new Dictionary
                    <QueryType, Func<IEnumerable<ExpandoObject>>>
                {
                    [QueryType.ComputersGroups] =
                        () => GetComputersGroupsData(),
                    [QueryType.ComputersSummaries] =
                        () => GetComputersSummariesData(),
                    [QueryType.DirectReportsDirectReports] =
                        () => GetUsersDirectReportsData(),
                    [QueryType.DirectReportsGroups] =
                        () => GetUserssGroupsData(),
                    [QueryType.DirectReportsSummaries] =
                        () => GetUsersSummariesData(),
                    [QueryType.GroupsComputers] =
                        () => GetGroupsComputersData(),
                    [QueryType.GroupsSummaries] =
                        () => GetGroupsSummariesData(),
                    [QueryType.GroupsUsers] = () => GetGroupsUsersData(),
                    [QueryType.GroupsUsersDirectReports] =
                        () => GetGroupsUsersDirectReportsData(),
                    [QueryType.GroupsUsersGroups] =
                        () => GetGroupsUsersGroupsData(),
                    [QueryType.OuComputers] = () => GetOuComputersData(
                        principalContext),
                    [QueryType.OuGroupsUsers] = () => GetOuGroupsUsersData(
                        principalContext)
                };
                data = mapping[_queryType]();
            }
            return data;
            // ReSharper restore AccessToDisposedClosure
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

        private static ComputerPrincipal FindComputerPrincipal(
            string distinguishedName)
        {
            using (var principalContext = GetPrincipalContext())
                return ComputerPrincipal.FindByIdentity(
                    principalContext,
                    IdentityType.DistinguishedName,
                    distinguishedName);
        }

        private static GroupPrincipal FindGroupPrincipal(
            string distinguishedName)
        {
            using (var principalContext = GetPrincipalContext())
                return GroupPrincipal.FindByIdentity(
                    principalContext,
                    IdentityType.DistinguishedName,
                    distinguishedName);
        }

        private static UserPrincipal FindUserPrincipal(
            string distinguishedName)
        {
            using (var principalContext = GetPrincipalContext())
                return UserPrincipal.FindByIdentity(
                    principalContext,
                    IdentityType.DistinguishedName,
                    distinguishedName);
        }

        private static PrincipalContext GetPrincipalContext()
        {
            return new PrincipalContext(ContextType.Domain);
        }

        private void AddAttributesToResult(
            dynamic result,
            ComputerPrincipal computerPrincipal = null,
            GroupPrincipal groupPrincipal = null,
            GroupPrincipal containerGroupPrincipal = null,
            UserPrincipal directReportUserPrincipal = null,
            UserPrincipal userPrincipal = null)
        {
            foreach (var property in _properties)
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

        private dynamic GenerateData(
            GroupPrincipal groupPrincipal,
            UserPrincipal userPrincipal,
            UserPrincipal directReportUserPrincipal)
        {
            dynamic data = new ExpandoObject();
            AddAttributesToResult(
                data,
                groupPrincipal: groupPrincipal,
                userPrincipal: userPrincipal,
                directReportUserPrincipal: directReportUserPrincipal);
            return data;
        }

        private dynamic GenerateData(
            GroupPrincipal groupPrincipal, ComputerPrincipal computerPrincipal)
        {
            dynamic data = new ExpandoObject();
            AddAttributesToResult(data, computerPrincipal, groupPrincipal);
            return data;
        }

        private dynamic GenerateData(UserPrincipal userPrincipal)
        {
            dynamic data = new ExpandoObject();
            AddAttributesToResult(data, userPrincipal: userPrincipal);
            return data;
        }

        private dynamic GenerateData(GroupPrincipal gp, UserPrincipal up)
        {
            dynamic data = new ExpandoObject();
            AddAttributesToResult(data, groupPrincipal: gp, userPrincipal: up);
            return data;
        }

        private dynamic GenerateData(UserPrincipal u, GroupPrincipal g)
        {
            dynamic data = new ExpandoObject();
            AddAttributesToResult(data, userPrincipal: u, groupPrincipal: g);
            return data;
        }

        private dynamic GenerateData(ComputerPrincipal cp, GroupPrincipal gp)
        {
            dynamic data = new ExpandoObject();
            AddAttributesToResult(
                data, computerPrincipal: cp, groupPrincipal: gp);
            return data;
        }

        private dynamic GenerateData(UserPrincipal up, UserPrincipal drup)
        {
            dynamic data = new ExpandoObject();
            AddAttributesToResult(
                data, userPrincipal: up, directReportUserPrincipal: drup);
            return data;
        }

        private dynamic GenerateData(GroupPrincipal gp)
        {
            dynamic data = new ExpandoObject();
            AddAttributesToResult(data, groupPrincipal: gp);
            return data;
        }

        private dynamic GenerateData(ComputerPrincipal cp)
        {
            dynamic result = new ExpandoObject();
            AddAttributesToResult(result, cp);
            return result;
        }

        private dynamic GenerateData(
            GroupPrincipal containerGroupPrincipal,
            UserPrincipal userPrincipal,
            GroupPrincipal groupPrincipal)
        {
            dynamic data = new ExpandoObject();
            AddAttributesToResult(
                data,
                containerGroupPrincipal: containerGroupPrincipal,
                userPrincipal: userPrincipal,
                groupPrincipal: groupPrincipal);
            return data;
        }

        private IEnumerable<ExpandoObject> GetComputersGroupsData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                using (var computerPrincipal = FindComputerPrincipal(
                    distinguishedName))
                {
                    if (computerPrincipal == null) continue;
                    using (var groups = computerPrincipal.GetGroups())
                    {
                        data.AddRange(
                            groups.GetGroupPrincipals().Select(
                                g => GenerateData(computerPrincipal, g))
                                .Cast<ExpandoObject>());
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetComputersSummariesData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                using (var computerPrincipal = FindComputerPrincipal(
                    distinguishedName))
                {
                    if (computerPrincipal == null) continue;
                    data.Add(GenerateData(computerPrincipal));
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetGroupsComputersData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                using (var groupPrincipal = FindGroupPrincipal(
                    distinguishedName))
                {
                    if (groupPrincipal == null) continue;
                    using (var members = groupPrincipal.GetMembers())
                    {
                        data.AddRange(
                            members.GetComputerPrincipals()
                                .Select(c => GenerateData(groupPrincipal, c))
                                .Cast<ExpandoObject>());
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetGroupsSummariesData()
        {
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                using (var groupPrincipal = FindGroupPrincipal(
                    distinguishedName))
                {
                    if (groupPrincipal == null) continue;
                    data.Add(GenerateData(groupPrincipal));
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetGroupsUsersData()
        {
            _properties = DefaultGroupUsersProperties;
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                using (var groupPrincipal = FindGroupPrincipal(
                    distinguishedName))
                {
                    if (groupPrincipal == null) continue;
                    using (var members = groupPrincipal.GetMembers())
                    {
                        data.AddRange(
                            members.GetUserPrincipals()
                                .Select(u => GenerateData(groupPrincipal, u))
                                .Cast<ExpandoObject>());
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetGroupsUsersDirectReportsData()
        {
            _properties = DefaultGroupUsersDirectReportsProperties;
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                using (var groupPrincipal = FindGroupPrincipal(
                    distinguishedName))
                {
                    if (groupPrincipal == null) continue;
                    using (var members = groupPrincipal.GetMembers())
                    {
                        foreach (var userPrincipal in
                            members.GetUserPrincipals())
                        {
                            foreach (var directReportDistinguishedName in
                                userPrincipal
                                    .GetDirectReportDistinguishedNames())
                            {
                                using (var directReportUserPrincipal =
                                    FindUserPrincipal(
                                        directReportDistinguishedName))
                                {
                                    data.Add(
                                        GenerateData(
                                            groupPrincipal,
                                            userPrincipal,
                                            directReportUserPrincipal));
                                }
                            }
                        }
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetGroupsUsersGroupsData()
        {
            _properties = DefaultGroupUsersGroupsProperties;
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                using (var groupPrincipal = FindGroupPrincipal(
                    distinguishedName))
                {
                    if (groupPrincipal == null) continue;
                    using (var members = groupPrincipal.GetMembers())
                    {
                        foreach (var userPrincipal in 
                            members.GetUserPrincipals())
                        {
                            using (var groups = userPrincipal.GetGroups())
                            {
                                data.AddRange(
                                    groups.GetGroupPrincipals()
                                        .Select(
                                            g => GenerateData(
                                                groupPrincipal,
                                                userPrincipal,
                                                g))
                                        .Cast<ExpandoObject>());
                            }
                        }
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetOuComputersData(
            PrincipalContext principalContext)
        {
            _properties = DefaultComputerProperties;
            var data = new List<ExpandoObject>();
            using (var computerPrincipal = new ComputerPrincipal(
                principalContext))
            {
                using (var principalSearcher = new PrincipalSearcher(
                    computerPrincipal))
                {
                    using (var principalSearchResult =
                        principalSearcher.FindAll())
                    {
                        data.AddRange(
                            principalSearchResult.GetComputerPrincipals()
                                .Select(c => GenerateData(c))
                                .Cast<ExpandoObject>());
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetOuGroupsUsersData(
            PrincipalContext principalContext)
        {
            _properties = DefaultGroupUsersProperties;
            var data = new List<ExpandoObject>();
            using (var principal = new GroupPrincipal(principalContext))
            {
                using (var principalSearcher = new PrincipalSearcher(
                    principal))
                {
                    using (var principalSearchResult =
                        principalSearcher.FindAll())
                    {
                        foreach (var groupPrincipal in
                            principalSearchResult.GetGroupPrincipals())
                        {
                            using (var users = groupPrincipal.GetMembers())
                            {
                                data.AddRange(
                                    users.GetUserPrincipals().Select(
                                        u => GenerateData(groupPrincipal, u))
                                        .Cast<ExpandoObject>());
                            }
                        }
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetUsersDirectReportsData()
        {
            _properties = DefaultUserDirectReportsProperties;
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                using (var userPrincipal = FindUserPrincipal(
                    distinguishedName))
                {
                    var directReportDistinguishedNames =
                        userPrincipal.GetDirectReportDistinguishedNames();
                    foreach (var directReportDistinguishedName in 
                        directReportDistinguishedNames)
                    {
                        using (var directReportUserPrincipal =
                            FindUserPrincipal(directReportDistinguishedName))
                        {
                            data.Add(
                                GenerateData(
                                    userPrincipal,
                                    directReportUserPrincipal));
                        }
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetUserssGroupsData()
        {
            _properties = DefaultUserGroupsProperties;
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                using (var userPrincipal = FindUserPrincipal(
                    distinguishedName))
                {
                    using (var groups = userPrincipal.GetGroups())
                    {
                        data.AddRange(
                            groups.GetGroupPrincipals()
                                .Select(g => GenerateData(userPrincipal, g))
                                .Cast<ExpandoObject>());
                    }
                }
            }
            return data;
        }

        private IEnumerable<ExpandoObject> GetUsersSummariesData()
        {
            _properties = DefaultUserProperties;
            var data = new List<ExpandoObject>();
            foreach (var distinguishedName in _distinguishedNames)
            {
                using (var userPrincipal = FindUserPrincipal(
                    distinguishedName))
                {
                    if (userPrincipal == null) continue;
                    data.Add(GenerateData(userPrincipal));
                }
            }
            return data;
        }
    }
}