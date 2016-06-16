using System;
using System.Collections.Generic;
using System.Data;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Linq;

namespace ActiveDirectoryToolWpf
{
    public enum ActiveDirectoryAttribute
    {
        Assistant,
        City,
        Comment,
        Company,
        ComputerGuid,
        ComputerHomeDirectory,
        ComputerHomeDrive,
        ComputerName,
        ComputerSid,
        ComputerUserPrincipalName,
        Country,
        Department,
        Division,
        EmailAddress,
        EmployeeId,
        Fax,
        GenerationalSuffix,
        GivenName,
        GroupGuid,
        GroupName,
        GroupSamAccountName,
        GroupSid,
        GroupUserPrincipalName,
        HomeAddress,
        HomePhone,
        Initials,
        IsAccountLockedOut,
        IsActive,
        ManagedBy,
        Manager,
        MiddleName,
        Mobile,
        Notes,
        Pager,
        ReversiblePasswordEncryption,
        Sip,
        State,
        StreetAddress,
        Surname,
        Title,
        UserAccountControl,
        UserAccountExpirationDate,
        UserAccountLockoutTime,
        UserAllowReversiblePasswordEncryption,
        UserBadLogonCount,
        UserCertificates,
        UserContext,
        UserContextType,
        UserDelegationPermitted,
        UserDescription,
        UserDisplayName,
        UserDistinguishedName,
        UserEnabled,
        UserGuid,
        UserHomeDirectory,
        UserHomeDrive,
        UserLastBadPasswordAttempt,
        UserLastLogon,
        UserLastPasswordSet,
        UserName,
        UserPasswordNeverExpires,
        UserPasswordNotRequired,
        UserPermittedLogonTimes,
        UserPermittedWorkstations,
        UserSamAccountName,
        UserScriptPath,
        UserSid,
        UserSmartcardLogonRequired,
        UserStructuralObjectClass,
        UserUserCannotChangePassword,
        UserUserPrincipalName,
        VoiceTelephoneNumber,
        Voip
    }

    internal static class Extensions
    {
        public static DataTable ToDataTable(this IEnumerable<dynamic> items)
        {
            var data = items.ToArray();
            if (!data.Any())
            {
                return null;
            }

            var dataTable = new DataTable();
            foreach (var key in ((IDictionary<string, object>) data[0]).Keys)
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

        internal List<ExpandoObject> GetResults()
        {
            var results = new List<ExpandoObject>();

            foreach (var data in Data)
            {
                GroupPrincipal groupPrincipal;
                UserPrincipal userPrincipal;
                if (data is UserGroups)
                {
                    var userGroups = data as UserGroups;
                    userPrincipal = userGroups.User;
                    foreach (var group in userGroups.Groups)
                    {
                        groupPrincipal = group;
                        dynamic result = new ExpandoObject();
                        AddPropertiesToResult(
                            null, groupPrincipal, null, userPrincipal, result);
                        results.Add(result);
                    }
                }
                else if (data is DirectReports)
                {
                    var directReports = data as DirectReports;
                    userPrincipal = directReports.User;
                    foreach (var directReport in directReports.Reports)
                    {
                        dynamic result = new ExpandoObject();
                        AddPropertiesToResult(
                            null, null, null, userPrincipal, result);
                        AddPropertiesToResult(
                            null, null, null, directReport, result);
                        results.Add(result);
                    }
                }
                else
                {
                    var computerPrincipal = data as ComputerPrincipal;
                    groupPrincipal = data as GroupPrincipal;
                    var principal = data as Principal;
                    userPrincipal = data as UserPrincipal;
                    dynamic result = new ExpandoObject();
                    AddPropertiesToResult(
                        computerPrincipal,
                        groupPrincipal,
                        principal,
                        userPrincipal,
                        result);
                    results.Add(result);
                }
            }
            return results;
        }

        private void AddPropertiesToResult(
            ComputerPrincipal computerPrincipal,
            GroupPrincipal groupPrincipal,
            Principal principal,
            UserPrincipal userPrincipal,
            dynamic result)
        {
            foreach (var attribute in Attributes)
            {
                switch (attribute)
                {
                    case ActiveDirectoryAttribute.UserAccountExpirationDate:
                        if (userPrincipal != null)
                        {
                            result.UserAccountExpirationDate =
                                userPrincipal.AccountExpirationDate;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserAccountLockoutTime:
                        if (userPrincipal != null)
                        {
                            result.UserAccountELockoutTime =
                                userPrincipal.AccountLockoutTime;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserAllowReversiblePasswordEncryption:
                        if (userPrincipal != null)
                        {
                            result.UserAllowReversiblePasswordEncryption =
                                userPrincipal
                                    .AllowReversiblePasswordEncryption;
                        }

                        break;

                    case ActiveDirectoryAttribute.Assistant:
                        if (userPrincipal != null)
                        {
                            result.Assistant =
                                userPrincipal.GetAssistant();
                        }

                        break;

                    case ActiveDirectoryAttribute.UserBadLogonCount:
                        if (userPrincipal != null)
                        {
                            result.UserBadLogonCount =
                                userPrincipal.BadLogonCount;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserCertificates:
                        if (userPrincipal != null)
                        {
                            result.UserCertificates =
                                userPrincipal.Certificates;
                        }

                        break;

                    case ActiveDirectoryAttribute.City:
                        if (userPrincipal != null)
                        {
                            result.City =
                                userPrincipal.GetCity();
                        }

                        break;

                    case ActiveDirectoryAttribute.Comment:
                        if (userPrincipal != null)
                        {
                            result.Comment =
                                userPrincipal.GetComment();
                        }

                        break;

                    case ActiveDirectoryAttribute.Company:
                        if (userPrincipal != null)
                        {
                            result.Company =
                                userPrincipal.GetCompany();
                        }

                        break;

                    case ActiveDirectoryAttribute.UserContext:
                        if (userPrincipal != null)
                        {
                            result.UserContext = userPrincipal.Context;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserContextType:
                        if (userPrincipal != null)
                        {
                            result.UserContextType = userPrincipal.ContextType;
                        }

                        break;

                    case ActiveDirectoryAttribute.Country:
                        if (userPrincipal != null)
                        {
                            result.Country =
                                userPrincipal.GetCountry();
                        }

                        break;

                    case ActiveDirectoryAttribute.UserDelegationPermitted:
                        if (userPrincipal != null)
                        {
                            result.UserDelegationPermitted =
                                userPrincipal.DelegationPermitted;
                        }

                        break;

                    case ActiveDirectoryAttribute.Department:
                        if (userPrincipal != null)
                        {
                            result.Department =
                                userPrincipal.GetDepartment();
                        }

                        break;

                    case ActiveDirectoryAttribute.UserDescription:
                        if (userPrincipal != null)
                        {
                            result.UserDescription = userPrincipal.Description;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserDisplayName:
                        if (userPrincipal != null)
                        {
                            result.UserDisplayName = userPrincipal.DisplayName;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserDistinguishedName:
                        if (userPrincipal != null)
                        {
                            result.UserDistinguishedName =
                                userPrincipal.DistinguishedName;
                        }

                        break;

                    case ActiveDirectoryAttribute.Division:
                        if (userPrincipal != null)
                        {
                            result.Division =
                                userPrincipal.GetDivision();
                        }

                        break;

                    case ActiveDirectoryAttribute.EmailAddress:
                        if (userPrincipal != null)
                        {
                            result.EmailAddress =
                                userPrincipal.EmailAddress;
                        }

                        break;

                    case ActiveDirectoryAttribute.EmployeeId:
                        if (userPrincipal != null)
                        {
                            result.EmployeeId =
                                userPrincipal.EmployeeId;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserEnabled:
                        if (userPrincipal != null)
                        {
                            result.UserEnabled =
                                userPrincipal.Enabled;
                        }

                        break;

                    case ActiveDirectoryAttribute.Fax:
                        if (userPrincipal != null)
                        {
                            result.Fax =
                                userPrincipal.GetFax();
                        }

                        break;

                    case ActiveDirectoryAttribute.GenerationalSuffix:
                        if (userPrincipal != null)
                        {
                            result.Suffix =
                                userPrincipal.GetSuffix();
                        }

                        break;

                    case ActiveDirectoryAttribute.GivenName:
                        if (userPrincipal != null)
                        {
                            result.GivenName =
                                userPrincipal.GivenName;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserGuid:
                        if (userPrincipal != null)
                        {
                            result.Guid = userPrincipal.Guid;
                        }

                        break;

                    case ActiveDirectoryAttribute.HomeAddress:
                        if (userPrincipal != null)
                        {
                            result.HomeAddress =
                                userPrincipal.GetHomeAddress();
                        }

                        break;

                    case ActiveDirectoryAttribute.UserHomeDirectory:
                        if (userPrincipal != null)
                        {
                            result.UserHomeDirectory =
                                userPrincipal.HomeDirectory;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserHomeDrive:
                        if (userPrincipal != null)
                        {
                            result.UserHomeDrive =
                                userPrincipal.HomeDrive;
                        }

                        break;

                    case ActiveDirectoryAttribute.HomePhone:
                        if (userPrincipal != null)
                        {
                            result.HomePhone =
                                userPrincipal.GetHomePhone();
                        }

                        break;

                    case ActiveDirectoryAttribute.Initials:
                        if (userPrincipal != null)
                        {
                            result.Initials =
                                userPrincipal.GetInitials();
                        }

                        break;

                    case ActiveDirectoryAttribute.IsActive:
                        if (userPrincipal != null)
                        {
                            result.IsActive =
                                userPrincipal.IsActive();
                        }

                        break;

                    case ActiveDirectoryAttribute.IsAccountLockedOut:
                        if (userPrincipal != null)
                        {
                            result.IsAccountLockedOut =
                                userPrincipal.IsAccountLockedOut();
                        }

                        break;

                    case ActiveDirectoryAttribute.UserLastBadPasswordAttempt:
                        if (userPrincipal != null)
                        {
                            result.UserLastBadPasswordAttempt =
                                userPrincipal
                                    .LastBadPasswordAttempt;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserLastLogon:
                        if (userPrincipal != null)
                        {
                            result.UserLastLogon =
                                userPrincipal.LastLogon;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserLastPasswordSet:
                        if (userPrincipal != null)
                        {
                            result.UserLastPasswordSet =
                                userPrincipal.LastPasswordSet;
                        }

                        break;

                    case ActiveDirectoryAttribute.ManagedBy:
                        if (groupPrincipal != null)
                        {
                            result.ManagedBy =
                                groupPrincipal.GetManagedBy();
                        }

                        break;

                    case ActiveDirectoryAttribute.Manager:
                        if (userPrincipal != null)
                        {
                            result.Manager =
                                userPrincipal.GetManager();
                        }

                        break;

                    case ActiveDirectoryAttribute.MiddleName:
                        if (userPrincipal != null)
                        {
                            result.MiddleName =
                                userPrincipal.MiddleName;
                        }

                        break;

                    case ActiveDirectoryAttribute.Mobile:
                        if (userPrincipal != null)
                        {
                            result.Mobile =
                                userPrincipal.GetMobile();
                        }

                        break;

                    case ActiveDirectoryAttribute.UserName:
                        if (userPrincipal != null)
                        {
                            result.UserName = userPrincipal.Name;
                        }

                        break;

                    case ActiveDirectoryAttribute.Notes:
                        if (userPrincipal != null)
                        {
                            result.Notes =
                                userPrincipal.GetNotes();
                        }

                        break;

                    case ActiveDirectoryAttribute.Pager:
                        if (userPrincipal != null)
                        {
                            result.Pager =
                                userPrincipal.GetPager();
                        }

                        break;

                    case ActiveDirectoryAttribute.UserPasswordNeverExpires:
                        if (userPrincipal != null)
                        {
                            result.UserPasswordNeverExpires =
                                userPrincipal
                                    .PasswordNeverExpires;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserPasswordNotRequired:
                        if (userPrincipal != null)
                        {
                            result.UserPasswordNotRequired =
                                userPrincipal.PasswordNotRequired;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserPermittedLogonTimes:
                        if (userPrincipal != null)
                        {
                            result.UserPermittedLogonTimes =
                                userPrincipal.PermittedLogonTimes;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserPermittedWorkstations:
                        if (userPrincipal != null)
                        {
                            result.UserPermittedWorkstations =
                                userPrincipal
                                    .PermittedWorkstations;
                        }

                        break;

                    case ActiveDirectoryAttribute
                        .ReversiblePasswordEncryption:
                        break;

                    case ActiveDirectoryAttribute.UserSamAccountName:
                        if (userPrincipal != null)
                        {
                            result.UserSamAccountName =
                                userPrincipal.SamAccountName;
                        }

                        break;

                    case ActiveDirectoryAttribute.GroupSamAccountName:
                        if (groupPrincipal != null)
                        {
                            result.GroupSamAccountName =
                                groupPrincipal.SamAccountName;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserScriptPath:
                        if (userPrincipal != null)
                        {
                            result.UserScriptPath =
                                userPrincipal.ScriptPath;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserSid:
                        if (userPrincipal != null)
                        {
                            result.UserSid = userPrincipal.Sid;
                        }

                        break;

                    case ActiveDirectoryAttribute.Sip:
                        if (userPrincipal != null)
                        {
                            result.Sip =
                                userPrincipal.GetSip();
                        }

                        break;

                    case ActiveDirectoryAttribute.UserSmartcardLogonRequired:
                        if (userPrincipal != null)
                        {
                            result.UserSmartcardLogonRequired =
                                userPrincipal
                                    .SmartcardLogonRequired;
                        }

                        break;

                    case ActiveDirectoryAttribute.State:
                        if (userPrincipal != null)
                        {
                            result.State =
                                userPrincipal.GetState();
                        }

                        break;

                    case ActiveDirectoryAttribute.StreetAddress:
                        if (userPrincipal != null)
                        {
                            result.Assistant =
                                userPrincipal.GetStreetAddress();
                        }

                        break;

                    case ActiveDirectoryAttribute.UserStructuralObjectClass:
                        if (userPrincipal != null)
                        {
                            result.UserStructuralObjectClass =
                                userPrincipal.StructuralObjectClass;
                        }

                        break;

                    case ActiveDirectoryAttribute.Surname:
                        if (userPrincipal != null)
                        {
                            result.Surname =
                                userPrincipal.Surname;
                        }

                        break;

                    case ActiveDirectoryAttribute.Title:
                        if (userPrincipal != null)
                        {
                            result.Title =
                                userPrincipal.GetTitle();
                        }

                        break;

                    case ActiveDirectoryAttribute.UserAccountControl:
                        if (userPrincipal != null)
                        {
                            result.UserAccountControl =
                                userPrincipal.GetUserAccountControl();
                        }

                        break;

                    case ActiveDirectoryAttribute.UserUserCannotChangePassword:
                        if (userPrincipal != null)
                        {
                            result.UserUserCannotChangePassword =
                                userPrincipal
                                    .UserCannotChangePassword;
                        }

                        break;

                    case ActiveDirectoryAttribute.UserUserPrincipalName:
                        if (userPrincipal != null)
                        {
                            result.UserUserPrincipalName =
                                userPrincipal.UserPrincipalName;
                        }

                        break;

                    case ActiveDirectoryAttribute.VoiceTelephoneNumber:
                        if (userPrincipal != null)
                        {
                            result.VoiceTelephoneNumber =
                                userPrincipal.VoiceTelephoneNumber;
                        }

                        break;

                    case ActiveDirectoryAttribute.Voip:
                        if (userPrincipal != null)
                        {
                            result.Voip =
                                userPrincipal.GetVoip();
                        }

                        break;

                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }
        }
    }
}