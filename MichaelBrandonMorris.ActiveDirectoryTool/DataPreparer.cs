using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Threading;
using MichaelBrandonMorris.Extensions.PrincipalExtensions;
using ActiveDirectoryPropertyMapping = System.Collections.Generic.Dictionary
    <MichaelBrandonMorris.ActiveDirectoryTool.ActiveDirectoryProperty, System.Action>;

namespace MichaelBrandonMorris.ActiveDirectoryTool
{
    internal class DataPreparer
    {
        private CancellationToken CancellationToken
        {
            get;
        }

        internal DataPreparer(CancellationToken cancellationToken)
        {
            CancellationToken = cancellationToken;
        }

        internal dynamic PrepareData(
            IEnumerable<ActiveDirectoryProperty> properties,
            ComputerPrincipal computerPrincipal = null,
            GroupPrincipal groupPrincipal = null,
            GroupPrincipal containerGroupPrincipal = null,
            UserPrincipal directReportUserPrincipal = null,
            UserPrincipal managerUserPrincipal = null,
            UserPrincipal userPrincipal = null)
        {
            dynamic result = new ExpandoObject();
            foreach (var property in properties)
            {
                CancellationToken.ThrowIfCancellationRequested();

                if (AddComputerProperty(
                    computerPrincipal, property, result))
                {
                    continue;
                }

                if (AddContainerGroupProperty(
                    containerGroupPrincipal, property, result))
                {
                    continue;
                }

                if (AddDirectReportProperty(
                    directReportUserPrincipal, property, result))
                {
                    continue;
                }

                if (AddGroupProperty(
                    groupPrincipal, property, result))
                {
                    continue;
                }

                AddUserProperty(userPrincipal, property, result);
            }

            return result;
        }

        private static bool AddComputerProperty(
            ComputerPrincipal computerPrincipal,
            ActiveDirectoryProperty property,
            dynamic result)
        {
            var propertyMapping = new ActiveDirectoryPropertyMapping
            {
                [ActiveDirectoryProperty.ComputerAccountExpirationDate] = () =>
                {
                    result.ComputerAccountExpirationDate =
                        computerPrincipal?.AccountExpirationDate;
                },
                [ActiveDirectoryProperty.ComputerAccountLockoutTime] = () =>
                {
                    result.ComputerAccountLockoutTime =
                        computerPrincipal?.AccountLockoutTime;
                },
                [ActiveDirectoryProperty.ComputerAllowReversiblePasswordEncryption] = () =>
                {
                    result.ComputerAllowReversiblePasswordEncryption =
                        computerPrincipal?
                            .AllowReversiblePasswordEncryption;
                },
                [ActiveDirectoryProperty.ComputerBadLogonCount] = () =>
                {
                    result.ComputerBadLogonCount =
                        computerPrincipal?.BadLogonCount;
                },
                [ActiveDirectoryProperty.ComputerCertificates] = () =>
                {
                    result.ComputerCertificates =
                        computerPrincipal?.Certificates;
                },
                [ActiveDirectoryProperty.ComputerContext] = () =>
                {
                    result.ComputerContext = computerPrincipal?.Context;
                },
                [ActiveDirectoryProperty.ComputerContextType] = () =>
                {
                    result.ComputerContextType =
                        computerPrincipal?.ContextType;
                },
                [ActiveDirectoryProperty.ComputerDelegationPermitted] = () =>
                {
                    result.ComputerDelegationPermitted =
                        computerPrincipal?.DelegationPermitted;
                },
                [ActiveDirectoryProperty.ComputerDescription] = () =>
                {
                    result.ComputerDescription =
                        computerPrincipal?.Description;
                },
                [ActiveDirectoryProperty.ComputerDisplayName] = () =>
                {
                    result.ComputerDisplayName =
                        computerPrincipal?.DisplayName;
                },
                [ActiveDirectoryProperty.ComputerDistinguishedName] = () =>
                {
                    result.ComputerDistinguishedName =
                        computerPrincipal?.DistinguishedName;
                },
                [ActiveDirectoryProperty.ComputerEnabled] = () =>
                {
                    result.ComputerEnabled = computerPrincipal?.Enabled;
                },
                [ActiveDirectoryProperty.ComputerGuid] = () =>
                {
                    result.ComputerGuid = computerPrincipal?.Guid;
                },
                [ActiveDirectoryProperty.ComputerHomeDirectory] = () =>
                {
                    result.ComputerHomeDirectory =
                        computerPrincipal?.HomeDirectory;
                },
                [ActiveDirectoryProperty.ComputerHomeDrive] = () =>
                {
                    result.ComputerHomeDrive =
                        computerPrincipal?.HomeDrive;
                },
                [ActiveDirectoryProperty.ComputerLastBadPasswordAttempt] = () =>
                {
                    result.ComputerLastBadPasswordAttempt =
                        computerPrincipal?.LastBadPasswordAttempt;
                },
                [ActiveDirectoryProperty.ComputerLastLogon] = () =>
                {
                    result.ComputerLastLogon =
                        computerPrincipal?.LastLogon;
                },
                [ActiveDirectoryProperty.ComputerLastPasswordSet] = () =>
                {
                    result.LastPasswordSet =
                        computerPrincipal?.LastPasswordSet;
                },
                [ActiveDirectoryProperty.ComputerName] = () =>
                {
                    result.ComputerName = computerPrincipal?.Name;
                },
                [ActiveDirectoryProperty.ComputerPasswordNeverExpires] = () =>
                {
                    result.ComputerPasswordNeverExpires =
                        computerPrincipal?.PasswordNeverExpires;
                },
                [ActiveDirectoryProperty.ComputerPasswordNotRequired] = () =>
                {
                    result.ComputerPasswordNotRequired =
                        computerPrincipal?.PasswordNotRequired;
                },
                [ActiveDirectoryProperty.ComputerPermittedLogonTimes] = () =>
                {
                    result.ComputerPermittedLogonTimes =
                        computerPrincipal?.PermittedLogonTimes;
                },
                [ActiveDirectoryProperty.ComputerPermittedWorkstations] = () =>
                {
                    result.ComputerPermittedWorkstations =
                        computerPrincipal?.PermittedWorkstations;
                },
                [ActiveDirectoryProperty.ComputerSamAccountName] = () =>
                {
                    result.ComputerSamAccountName =
                        computerPrincipal?.SamAccountName;
                },
                [ActiveDirectoryProperty.ComputerScriptPath] = () =>
                {
                    result.ComputerScriptPath =
                        computerPrincipal?.ScriptPath;
                },
                [ActiveDirectoryProperty.ComputerServicePrincipalNames] = () =>
                {
                    result.ComputerServicecomputerPrincipalNames =
                        computerPrincipal?.ServicePrincipalNames;
                },
                [ActiveDirectoryProperty.ComputerSid] = () =>
                {
                    result.ComputerSid = computerPrincipal?.Sid;
                },
                [ActiveDirectoryProperty.ComputerSmartcardLogonRequired] = () =>
                {
                    result.ComputerSmartcardLogonRequired =
                        computerPrincipal?.SmartcardLogonRequired;
                },
                [ActiveDirectoryProperty.ComputerStructuralObjectClass] = () =>
                {
                    result.ComputerStructuralObjectClass =
                        computerPrincipal?.StructuralObjectClass;
                },
                [ActiveDirectoryProperty.ComputerUserCannotChangePassword] = () =>
                {
                    result.ComputerUserCannotChangePassword =
                        computerPrincipal?.UserCannotChangePassword;
                },
                [ActiveDirectoryProperty.ComputerUserPrincipalName] = () =>
                {
                    result.ComputerUsercomputerPrincipalName =
                        computerPrincipal?.UserPrincipalName;
                }
            };

            if (!propertyMapping.ContainsKey(property)) return false;
            propertyMapping[property]();
            return true;
        }

        private static bool AddContainerGroupProperty(
            GroupPrincipal containerGroupPrincipal,
            ActiveDirectoryProperty property,
            dynamic result)
        {
            var propertyMapping = new ActiveDirectoryPropertyMapping
            {
                [ActiveDirectoryProperty.ContainerGroupContext] = () =>
                {
                    result.ContainerGroupContext =
                        containerGroupPrincipal?.Context;
                },
                [ActiveDirectoryProperty.ContainerGroupContextType] = () =>
                {
                    result.ContainerGroupContextType =
                        containerGroupPrincipal?.ContextType;
                },
                [ActiveDirectoryProperty.ContainerGroupDescription] = () =>
                {
                    result.ContainerGroupDescription =
                        containerGroupPrincipal?.Description;
                },
                [ActiveDirectoryProperty.ContainerGroupDisplayName] = () =>
                {
                    result.ContainerGroupDisplayName =
                        containerGroupPrincipal?.DisplayName;
                },
                [ActiveDirectoryProperty.ContainerGroupDistinguishedName] = () =>
                {
                    result.ContainerGroupDistinguishedName =
                        containerGroupPrincipal?.DistinguishedName;
                },
                [ActiveDirectoryProperty.ContainerGroupGuid] = () =>
                {
                    result.ContainerGroupGuid = containerGroupPrincipal?.Guid;
                },
                [ActiveDirectoryProperty.ContainerGroupIsSecurityGroup] = () =>
                {
                    result.ContainerGroupIsSecurityGroup =
                        containerGroupPrincipal?.IsSecurityGroup;
                },
                [ActiveDirectoryProperty.ContainerGroupManagedByDistinguishedName] = () =>
                {
                    result.ContainerGroupManagedByDistinguishedName =
                        containerGroupPrincipal?.GetManagedByDistinguishedName();
                },
                [ActiveDirectoryProperty.ContainerGroupManagedByName] = () =>
                {
                    result.ContainerGroupManagedByName =
                        containerGroupPrincipal.GetManagedByName();
                },
                [ActiveDirectoryProperty.ContainerGroupName] = () =>
                {
                    result.ContainerGroupName = containerGroupPrincipal?.Name;
                },
                [ActiveDirectoryProperty.ContainerGroupSamAccountName] = () =>
                {
                    result.ContainerGroupSamAccountName =
                        containerGroupPrincipal?.SamAccountName;
                },
                [ActiveDirectoryProperty.ContainerGroupScope] = () =>
                {
                    result.ContainerGroupScope =
                        containerGroupPrincipal?.GroupScope;
                },
                [ActiveDirectoryProperty.ContainerGroupSid] = () =>
                {
                    result.ContainerGroupSid = containerGroupPrincipal?.Sid;
                },
                [ActiveDirectoryProperty.ContainerGroupStructuralObjectClass] = () =>
                {
                    result.ContainerGroupStructuralObjectClass =
                        containerGroupPrincipal?.StructuralObjectClass;
                },
                [ActiveDirectoryProperty.ContainerGroupUserPrincipalName] = () =>
                {
                    result.ContainerGroupUserPrincipalName =
                        containerGroupPrincipal?.UserPrincipalName;
                },
                [ActiveDirectoryProperty.ContainerGroupMembers] = () =>
                {
                    result.ContainerGroupMembers =
                        containerGroupPrincipal?.Members;
                }
            };

            if (!propertyMapping.ContainsKey(property)) return false;
            propertyMapping[property]();
            return true;
        }

        private static bool AddDirectReportProperty(
            UserPrincipal directReportUserPrincipal,
            ActiveDirectoryProperty property,
            dynamic result)
        {
            var propertyMapping = new ActiveDirectoryPropertyMapping
            {
                [ActiveDirectoryProperty.DirectReportUserAccountControl] = () =>
                {
                    result.DirectReportAccountControl =
                        directReportUserPrincipal.GetUserAccountControl();
                },
                [ActiveDirectoryProperty.DirectReportAccountExpirationDate] = () =>
                {
                    result.DirectReportAccountExpirationDate =
                        directReportUserPrincipal.AccountExpirationDate;
                },
                [ActiveDirectoryProperty.DirectReportAccountLockoutTime] = () =>
                {
                    result.DirectReportAccountLockoutTime =
                        directReportUserPrincipal.AccountLockoutTime;
                },
                [ActiveDirectoryProperty.DirectReportAllowReversiblePasswordEncryption] = () =>
                {
                    result.DirectReportAllowReversiblePasswordEncryption =
                        directReportUserPrincipal
                            .AllowReversiblePasswordEncryption;
                },
                [ActiveDirectoryProperty.DirectReportAssistant] = () =>
                {
                    result.DirectReportAssistant =
                        directReportUserPrincipal.GetAssistant();
                },
                [ActiveDirectoryProperty.DirectReportBadLogonCount] = () =>
                {
                    result.DirectReportBadLogonCount =
                        directReportUserPrincipal.BadLogonCount;
                },
                [ActiveDirectoryProperty.DirectReportCertificates] = () =>
                {
                    result.DirectReportCertificates =
                        directReportUserPrincipal.Certificates;
                },
                [ActiveDirectoryProperty.DirectReportCity] = () =>
                {
                    result.DirectReportCity =
                        directReportUserPrincipal.GetCity();
                },
                [ActiveDirectoryProperty.DirectReportComment] = () =>
                {
                    result.DirectReportComment =
                        directReportUserPrincipal.GetComment();
                },
                [ActiveDirectoryProperty.DirectReportCompany] = () =>
                {
                    result.DirectReportCompany =
                        directReportUserPrincipal.GetCompany();
                },
                [ActiveDirectoryProperty.DirectReportContext] = () =>
                {
                    result.DirectReportContext =
                        directReportUserPrincipal.Context;
                },
                [ActiveDirectoryProperty.DirectReportContextType] = () =>
                {
                    result.DirectReportContextType =
                        directReportUserPrincipal.ContextType;
                },
                [ActiveDirectoryProperty.DirectReportCountry] = () =>
                {
                    result.DirectReportCountry =
                        directReportUserPrincipal.GetCountry();
                },
                [ActiveDirectoryProperty.DirectReportDelegationPermitted] = () =>
                {
                    result.DirectReportDelegationPermitted =
                        directReportUserPrincipal.DelegationPermitted;
                },
                [ActiveDirectoryProperty.DirectReportDepartment] = () =>
                {
                    result.DirectReportDepartment =
                        directReportUserPrincipal.GetDepartment();
                },
                [ActiveDirectoryProperty.DirectReportDescription] = () =>
                {
                    result.DirectReportDescription =
                        directReportUserPrincipal.Description;
                },
                [ActiveDirectoryProperty.DirectReportDisplayName] = () =>
                {
                    result.DirectReportDisplayName =
                        directReportUserPrincipal.DisplayName;
                },
                [ActiveDirectoryProperty.DirectReportDistinguishedName] = () =>
                {
                    result.DirectReportDistinguishedName =
                        directReportUserPrincipal?.DistinguishedName;
                },
                [ActiveDirectoryProperty.DirectReportDivision] = () =>
                {
                    result.DirectReportDivision =
                        directReportUserPrincipal.GetDivision();
                },
                [ActiveDirectoryProperty.DirectReportEmailAddress] = () =>
                {
                    result.DirectReportEmailAddress =
                        directReportUserPrincipal.EmailAddress;
                },
                [ActiveDirectoryProperty.DirectReportEmployeeId] = () =>
                {
                    result.DirectReportEmployeeId =
                        directReportUserPrincipal.EmployeeId;
                },
                [ActiveDirectoryProperty.DirectReportEnabled] = () =>
                {
                    result.DirectReportEnabled =
                        directReportUserPrincipal.Enabled;
                },
                [ActiveDirectoryProperty.DirectReportFax] = () =>
                {
                    result.DirectReportFax =
                        directReportUserPrincipal.GetFax();
                },
                [ActiveDirectoryProperty.DirectReportSuffix] = () =>
                {
                    result.DirectReportSuffix =
                        directReportUserPrincipal.GetSuffix();
                },
                [ActiveDirectoryProperty.DirectReportGivenName] = () =>
                {
                    result.DirectReportGivenName =
                        directReportUserPrincipal.GivenName;
                },
                [ActiveDirectoryProperty.DirectReportGuid] = () =>
                {
                    result.DirectReportGuid =
                        directReportUserPrincipal.Guid;
                },
                [ActiveDirectoryProperty.DirectReportHomeAddress] = () =>
                {
                    result.DirectReportHomeAddress =
                        directReportUserPrincipal.GetHomeAddress();
                },
                [ActiveDirectoryProperty.DirectReportHomeDirectory] = () =>
                {
                    result.DirectReportHomeDirectory =
                        directReportUserPrincipal.HomeDirectory;
                },
                [ActiveDirectoryProperty.DirectReportHomeDrive] = () =>
                {
                    result.DirectReportHomeDrive =
                        directReportUserPrincipal.HomeDrive;
                },
                [ActiveDirectoryProperty.DirectReportHomePhone] = () =>
                {
                    result.DirectReportHomePhone =
                        directReportUserPrincipal.GetHomePhone();
                },
                [ActiveDirectoryProperty.DirectReportInitials] = () =>
                {
                    result.DirectReportInitials =
                        directReportUserPrincipal.GetInitials();
                },
                [ActiveDirectoryProperty.DirectReportIsAccountLockedOut] = () =>
                {
                    result.DirectReportIsAccountLockedOut =
                        directReportUserPrincipal.IsAccountLockedOut();
                },
                [ActiveDirectoryProperty.DirectReportIsActive] = () =>
                {
                    result.DirectReportIsActive =
                        directReportUserPrincipal.IsActive();
                },
                [ActiveDirectoryProperty.DirectReportLastBadPasswordAttempt] = () =>
                {
                    result.DirectReportLastBadPasswordAttempt =
                        directReportUserPrincipal.LastBadPasswordAttempt;
                },
                [ActiveDirectoryProperty.DirectReportLastLogon] = () =>
                {
                    result.DirectReportLastLogon =
                        directReportUserPrincipal.LastLogon;
                },
                [ActiveDirectoryProperty.DirectReportLastPasswordSet] = () =>
                {
                    result.DirectReportLastPasswordSet =
                        directReportUserPrincipal.LastPasswordSet;
                },
                [ActiveDirectoryProperty.DirectReportMiddleName] = () =>
                {
                    result.DirectReportMiddleName =
                        directReportUserPrincipal.MiddleName;
                },
                [ActiveDirectoryProperty.DirectReportMobile] = () =>
                {
                    result.DirectReportMobile =
                        directReportUserPrincipal.GetMobile();
                },
                [ActiveDirectoryProperty.DirectReportName] = () =>
                {
                    result.DirectReportName =
                        directReportUserPrincipal?.Name;
                },
                [ActiveDirectoryProperty.DirectReportNotes] = () =>
                {
                    result.DirectReportNotes =
                        directReportUserPrincipal.GetNotes();
                },
                [ActiveDirectoryProperty.DirectReportPager] = () =>
                {
                    result.DirectReportPager =
                        directReportUserPrincipal.GetPager();
                },
                [ActiveDirectoryProperty.DirectReportPasswordNeverExpires] = () =>
                {
                    result.DirectReportPasswordNeverExpires =
                        directReportUserPrincipal.PasswordNeverExpires;
                },
                [ActiveDirectoryProperty.DirectReportPasswordNotRequired] = () =>
                {
                    result.DirectReportPasswordNotRequired =
                        directReportUserPrincipal.PasswordNotRequired;
                },
                [ActiveDirectoryProperty.DirectReportPermittedLogonTimes] = () =>
                {
                    result.DirectReportPermittedLogonTimes =
                        directReportUserPrincipal.PermittedLogonTimes;
                },
                [ActiveDirectoryProperty.DirectReportPermittedWorkstations] = () =>
                {
                    result.DirectReportPermittedWorkstations =
                        directReportUserPrincipal.PermittedWorkstations;
                },
                [ActiveDirectoryProperty.DirectReportSamAccountName] = () =>
                {
                    result.DirectReportSamAccountName =
                        directReportUserPrincipal.SamAccountName;
                },
                [ActiveDirectoryProperty.DirectReportScriptPath] = () =>
                {
                    result.DirectReportScriptPath =
                        directReportUserPrincipal.ScriptPath;
                },
                [ActiveDirectoryProperty.DirectReportSid] = () =>
                {
                    result.DirectReportSid = directReportUserPrincipal.Sid;
                },
                [ActiveDirectoryProperty.DirectReportSip] = () =>
                {
                    result.DirectReportSip =
                        directReportUserPrincipal.GetSip();
                },
                [ActiveDirectoryProperty.DirectReportSmartcardLogonRequired] = () =>
                {
                    result.DirectReportSmartcardLogonRequired =
                        directReportUserPrincipal.SmartcardLogonRequired;
                },
                [ActiveDirectoryProperty.DirectReportState] = () =>
                {
                    result.DirectReportState =
                        directReportUserPrincipal.GetState();
                },
                [ActiveDirectoryProperty.DirectReportStreetAddress] = () =>
                {
                    result.DirectReportStreetAddress =
                        directReportUserPrincipal.GetStreetAddress();
                },
                [ActiveDirectoryProperty.DirectReportStructuralObjectClass] = () =>
                {
                    result.DirectReportStructuralObjectClass =
                        directReportUserPrincipal.StructuralObjectClass;
                },
                [ActiveDirectoryProperty.DirectReportSurname] = () =>
                {
                    result.DirectReportSurname =
                        directReportUserPrincipal.Surname;
                },
                [ActiveDirectoryProperty.DirectReportTitle] = () =>
                {
                    result.DirectReportTitle =
                        directReportUserPrincipal.GetTitle();
                },
                [ActiveDirectoryProperty.DirectReportUserCannotChangePassword] = () =>
                {
                    result.DirectReportUserCannotChangePassword =
                        directReportUserPrincipal.UserCannotChangePassword;
                },
                [ActiveDirectoryProperty.DirectReportUserPrincipalName] = () =>
                {
                    result.DirectReportdirectReportUserPrincipalName =
                        directReportUserPrincipal.UserPrincipalName;
                },
                [ActiveDirectoryProperty.DirectReportVoiceTelephoneNumber] = () =>
                {
                    result.DirectReportVoiceTelephoneNumber =
                        directReportUserPrincipal.VoiceTelephoneNumber;
                },
                [ActiveDirectoryProperty.DirectReportVoip] = () =>
                {
                    result.DirectReportVoip =
                        directReportUserPrincipal.GetVoip();
                }
            };

            if (!propertyMapping.ContainsKey(property)) return false;
            propertyMapping[property]();
            return true;
        }

        private static bool AddGroupProperty(
            GroupPrincipal groupPrincipal,
            ActiveDirectoryProperty property,
            dynamic result)
        {
            var propertyMapping = new ActiveDirectoryPropertyMapping
            {
                [ActiveDirectoryProperty.GroupContext] = () =>
                    {
                        result.GroupContext = groupPrincipal.Context;
                    },
                [ActiveDirectoryProperty.GroupContextType] = () =>
                {
                    result.GroupContextType = groupPrincipal.ContextType;
                },
                [ActiveDirectoryProperty.GroupDescription] = () =>
                {
                    result.GroupDescription = groupPrincipal.Description;
                },
                [ActiveDirectoryProperty.GroupDisplayName] = () =>
                {
                    result.GroupDisplayName = groupPrincipal.DisplayName;
                },
                [ActiveDirectoryProperty.GroupDistinguishedName] = () =>
                {
                    result.GroupDistinguishedName =
                        groupPrincipal.DistinguishedName;
                },
                [ActiveDirectoryProperty.GroupGuid] = () =>
                {
                    result.GroupGuid = groupPrincipal.Guid;
                },
                [ActiveDirectoryProperty.GroupIsSecurityGroup] = () =>
                {
                    result.GroupIsSecurityGroup =
                        groupPrincipal.IsSecurityGroup;
                },
                [ActiveDirectoryProperty.GroupManagedByDistinguishedName] = () =>
                {
                    result.GroupManagedByDistinguishedName = 
                        groupPrincipal.GetManagedByDistinguishedName();
                },
                [ActiveDirectoryProperty.GroupManagedByName] = () =>
                {
                    result.GroupManagedByName =
                        groupPrincipal.GetManagedByName();
                },
                [ActiveDirectoryProperty.GroupName] = () =>
                {
                    result.GroupName = groupPrincipal.Name;
                },
                [ActiveDirectoryProperty.GroupSamAccountName] = () =>
                {
                    result.GroupSamAccountName = groupPrincipal.SamAccountName;
                },
                [ActiveDirectoryProperty.GroupScope] = () =>
                {
                    result.GroupScope = groupPrincipal.GroupScope;
                },
                [ActiveDirectoryProperty.GroupSid] = () =>
                {
                    result.GroupSid = groupPrincipal.Sid;
                },
                [ActiveDirectoryProperty.GroupStructuralObjectClass] = () =>
                {
                    result.GroupStructuralObjectClass =
                        groupPrincipal.StructuralObjectClass;
                },
                [ActiveDirectoryProperty.GroupUserPrincipalName] = () =>
                {
                    result.GroupUserPrincipalName =
                        groupPrincipal.UserPrincipalName;
                },
                [ActiveDirectoryProperty.GroupMembers] = () =>
                {
                    result.GroupMembers = groupPrincipal.Members;
                }
            };

            if (!propertyMapping.ContainsKey(property)) return false;
            propertyMapping[property]();
            return true;
        }

        private static void AddUserProperty(
            UserPrincipal userPrincipal,
            ActiveDirectoryProperty property,
            dynamic result)
        {
            var propertyMapping = new ActiveDirectoryPropertyMapping
            {
                [ActiveDirectoryProperty.UserUserAccountControl] = () =>
                {
                    result.UserAccountControl =
                        userPrincipal.GetUserAccountControl();
                },
                [ActiveDirectoryProperty.UserAccountExpirationDate] = () =>
                {
                    result.UserAccountExpirationDate =
                        userPrincipal.AccountExpirationDate;
                },
                [ActiveDirectoryProperty.UserAccountLockoutTime] = () =>
                {
                    result.UserAccountLockoutTime =
                        userPrincipal.AccountLockoutTime;
                },
                [ActiveDirectoryProperty.UserAllowReversiblePasswordEncryption] = () =>
                {
                    result.UserAllowReversiblePasswordEncryption =
                        userPrincipal.AllowReversiblePasswordEncryption;
                },
                [ActiveDirectoryProperty.UserAssistant] = () =>
                {
                    result.UserAssistant = userPrincipal.GetAssistant();
                },
                [ActiveDirectoryProperty.UserBadLogonCount] = () =>
                {
                    result.UserBadLogonCount = userPrincipal.BadLogonCount;
                },
                [ActiveDirectoryProperty.UserCertificates] = () =>
                {
                    result.UserCertificates = userPrincipal.Certificates;
                },
                [ActiveDirectoryProperty.UserCity] = () =>
                {
                    result.UserCity = userPrincipal.GetCity();
                },
                [ActiveDirectoryProperty.UserComment] = () =>
                {
                    result.UserComment = userPrincipal.GetComment();
                },
                [ActiveDirectoryProperty.UserCompany] = () =>
                {
                    result.UserCompany = userPrincipal.GetCompany();
                },
                [ActiveDirectoryProperty.UserContext] = () =>
                {
                    result.UserContext = userPrincipal.Context;
                },
                [ActiveDirectoryProperty.UserContextType] = () =>
                {
                    result.UserContextType = userPrincipal.ContextType;
                },
                [ActiveDirectoryProperty.UserCountry] = () =>
                {
                    result.UserCountry = userPrincipal.GetCountry();
                },
                [ActiveDirectoryProperty.UserDelegationPermitted] = () =>
                {
                    result.UserDelegationPermitted =
                        userPrincipal.DelegationPermitted;
                },
                [ActiveDirectoryProperty.UserDepartment] = () =>
                {
                    result.UserDepartment = userPrincipal.GetDepartment();
                },
                [ActiveDirectoryProperty.UserDescription] = () =>
                {
                    result.UserDescription = userPrincipal.Description;
                },
                [ActiveDirectoryProperty.UserDisplayName] = () =>
                {
                    result.UserDisplayName = userPrincipal.DisplayName;
                },
                [ActiveDirectoryProperty.UserDistinguishedName] = () =>
                {
                    result.UserDistinguishedName =
                        userPrincipal.DistinguishedName;
                },
                [ActiveDirectoryProperty.UserDivision] = () =>
                {
                    result.UserDivision = userPrincipal.GetDivision();
                },
                [ActiveDirectoryProperty.UserEmailAddress] = () =>
                {
                    result.UserEmailAddress = userPrincipal.EmailAddress;
                },
                [ActiveDirectoryProperty.UserEmployeeId] = () =>
                {
                    result.UserEmployeeId = userPrincipal.EmployeeId;
                },
                [ActiveDirectoryProperty.UserEnabled] = () =>
                {
                    result.UserEnabled = userPrincipal.Enabled;
                },
                [ActiveDirectoryProperty.UserFax] = () =>
                {
                    result.UserFax = userPrincipal.GetFax();
                },
                [ActiveDirectoryProperty.UserSuffix] = () =>
                {
                    result.UserSuffix = userPrincipal.GetSuffix();
                },
                [ActiveDirectoryProperty.UserGivenName] = () =>
                {
                    result.UserGivenName = userPrincipal.GivenName;
                },
                [ActiveDirectoryProperty.UserGuid] = () =>
                {
                    result.UserGuid = userPrincipal.Guid;
                },
                [ActiveDirectoryProperty.UserHomeAddress] = () =>
                {
                    result.UserHomeAddress = userPrincipal.GetHomeAddress();
                },
                [ActiveDirectoryProperty.UserHomeDirectory] = () =>
                {
                    result.UserHomeDirectory = userPrincipal.HomeDirectory;
                },
                [ActiveDirectoryProperty.UserHomeDrive] = () =>
                {
                    result.UserHomeDrive = userPrincipal.HomeDrive;
                },
                [ActiveDirectoryProperty.UserHomePhone] = () =>
                {
                    result.UserHomePhone = userPrincipal.GetHomePhone();
                },
                [ActiveDirectoryProperty.UserInitials] = () =>
                {
                    result.UserInitials = userPrincipal.GetInitials();
                },
                [ActiveDirectoryProperty.UserIsAccountLockedOut] = () =>
                {
                    result.UserIsAccountLockedOut =
                        userPrincipal.IsAccountLockedOut();
                },
                [ActiveDirectoryProperty.UserIsActive] = () =>
                {
                    result.UserIsActive = userPrincipal.IsActive();
                },
                [ActiveDirectoryProperty.UserLastBadPasswordAttempt] = () =>
                {
                    result.UserLastBadPasswordAttempt =
                        userPrincipal.LastBadPasswordAttempt;
                },
                [ActiveDirectoryProperty.UserLastLogon] = () =>
                {
                    result.UserLastLogon = userPrincipal.LastLogon;
                },
                [ActiveDirectoryProperty.UserLastPasswordSet] = () =>
                {
                    result.UserLastPasswordSet = userPrincipal.LastPasswordSet;
                },
                [ActiveDirectoryProperty.ManagerDistinguishedName] = () =>
                {
                    result.ManagerDistinguishedName = userPrincipal
                        .GetManagerDistinguishedName();
                },
                [ActiveDirectoryProperty.ManagerName] = () =>
                {
                    result.ManagerName = userPrincipal.GetManagerName();
                },
                [ActiveDirectoryProperty.UserMiddleName] = () =>
                {
                    result.UserMiddleName = userPrincipal.MiddleName;
                },
                [ActiveDirectoryProperty.UserMobile] = () =>
                {
                    result.UserMobile = userPrincipal.GetMobile();
                },
                [ActiveDirectoryProperty.UserName] = () =>
                {
                    result.UserName = userPrincipal.Name;
                },
                [ActiveDirectoryProperty.UserNotes] = () =>
                {
                    result.UserNotes = userPrincipal.GetNotes();
                },
                [ActiveDirectoryProperty.UserPager] = () =>
                {
                    result.UserPager = userPrincipal.GetPager();
                },
                [ActiveDirectoryProperty.UserPasswordNeverExpires] = () =>
                {
                    result.UserPasswordNeverExpires =
                        userPrincipal.PasswordNeverExpires;
                },
                [ActiveDirectoryProperty.UserPasswordNotRequired] = () =>
                {
                    result.UserPasswordNotRequired =
                        userPrincipal.PasswordNotRequired;
                },
                [ActiveDirectoryProperty.UserPermittedLogonTimes] = () =>
                {
                    result.UserPermittedLogonTimes =
                        userPrincipal.PermittedLogonTimes;
                },
                [ActiveDirectoryProperty.UserPermittedWorkstations] = () =>
                {
                    result.UserPermittedWorkstations =
                        userPrincipal.PermittedWorkstations;
                },
                [ActiveDirectoryProperty.UserSamAccountName] = () =>
                {
                    result.UserSamAccountName = userPrincipal.SamAccountName;
                },
                [ActiveDirectoryProperty.UserScriptPath] = () =>
                {
                    result.UserScriptPath = userPrincipal.ScriptPath;
                },
                [ActiveDirectoryProperty.UserSid] = () =>
                {
                    result.UserSid = userPrincipal.Sid;
                },
                [ActiveDirectoryProperty.UserSip] = () =>
                {
                    result.UserSip = userPrincipal.GetSip();
                },
                [ActiveDirectoryProperty.UserSmartcardLogonRequired] = () =>
                {
                    result.UserSmartcardLogonRequired =
                        userPrincipal.SmartcardLogonRequired;
                },
                [ActiveDirectoryProperty.UserState] = () =>
                {
                    result.UserState = userPrincipal.GetState();
                },
                [ActiveDirectoryProperty.UserStreetAddress] = () =>
                {
                    result.UserStreetAddress =
                        userPrincipal.GetStreetAddress();
                },
                [ActiveDirectoryProperty.UserStructuralObjectClass] = () =>
                {
                    result.UserStructuralObjectClass =
                        userPrincipal.StructuralObjectClass;
                },
                [ActiveDirectoryProperty.UserSurname] = () =>
                {
                    result.UserSurname = userPrincipal.Surname;
                },
                [ActiveDirectoryProperty.UserTitle] = () =>
                {
                    result.UserTitle = userPrincipal.GetTitle();
                },
                [ActiveDirectoryProperty.UserUserCannotChangePassword] = () =>
                {
                    result.UserUserCannotChangePassword =
                        userPrincipal.UserCannotChangePassword;
                },
                [ActiveDirectoryProperty.UserUserPrincipalName] = () =>
                {
                    result.UserUserPrincipalName =
                        userPrincipal.UserPrincipalName;
                },
                [ActiveDirectoryProperty.UserVoiceTelephoneNumber] = () =>
                {
                    result.UserVoiceTelephoneNumber =
                        userPrincipal.VoiceTelephoneNumber;
                },
                [ActiveDirectoryProperty.UserVoip] = () =>
                {
                    result.UserVoip = userPrincipal.GetVoip();
                }
            };

            if (!propertyMapping.ContainsKey(property)) return;
            propertyMapping[property]();
        }
    }
}
