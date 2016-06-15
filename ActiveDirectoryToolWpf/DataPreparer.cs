using System;
using System.Collections.Generic;
using System.Data;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiveDirectoryToolWpf
{
    internal class DataPreparer
    {
        internal List<ActiveDirectoryUserAttribute> Attributes { get; set; }
        internal IEnumerable<Principal> Principals { get; set; }

        internal List<ExpandoObject> GetResults()
        {
            var results = new List<ExpandoObject>();

            foreach (var principal in Principals)
            {
                dynamic result = new ExpandoObject();
                result.Name = principal.Name;
                result.DistinguishedName = principal.DistinguishedName;
                results.Add(result);
            }
            return results;
        }
    }

    internal static class Extensions
    {
        public static DataTable ToDataTable(this IEnumerable<dynamic> items)
        {
            var data = items.ToArray();
            if (data.Count() == 0) return null;

            var dt = new DataTable();
            foreach (var key in ((IDictionary<string, object>)data[0]).Keys)
            {
                dt.Columns.Add(key);
            }
            foreach (var d in data)
            {
                dt.Rows.Add(((IDictionary<string, object>)d).Values.ToArray());
            }
            return dt;
        }
    }

    internal enum ActiveDirectoryUserAttribute
    {
        AccountExpirationDate,
        AccountLockoutTime,
        Assistant,
        BadLogonCount,
        Certificates,
        City,
        Comment,
        Company,
        Country,
        DelegationPermitted,
        Department,
        Description,
        DisplayName,
        DistinguishedName,
        Division,
        EmailAddress,
        EmployeeId,
        Enabled,
        Fax,
        GenerationalSuffix,
        GivenName,
        Guid,
        HomeAddress,
        HomeDirectory,
        HomeDrive,
        HomePhone,
        Initials,
        LastBadPasswordAttempt,
        LastLogon,
        LastPasswordSet,
        Manager,
        MiddleName,
        Mobile,
        Name,
        Notes,
        Pager,
        PasswordNeverExpires,
        PasswordNotRequired,
        PermittedLogonTimes,
        PermittedWorkstations,
        ReversiblePasswordEncryption,
        SamAccountName,
        ScriptPath,
        Sid,
        Sip,
        SmartcardLogonRequired,
        State,
        StreetAddress,
        Surname,
        Title,
        UserAccountControl,
        UserCannotChangePassword,
        UserPrincipalName,
        VoiceTelephoneNumber,
        Voip
    }
}
