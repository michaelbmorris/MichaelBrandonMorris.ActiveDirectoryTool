using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.Dynamic;
using System.Threading.Tasks;

namespace ActiveDirectoryToolWpf
{
    public enum QueryType
    {
        ContextComputerGroups,
        ContextComputerSummary,
        ContextDirectReportDirectReports,
        ContextDirectReportGroups,
        ContextDirectReportSummary,
        ContextGroupComputers,
        ContextGroupUsers,
        ContextGroupUsersDirectReports,
        ContextGroupUsersGroups,
        ContextGroupSummary,
        ContextUserDirectReports,
        ContextUserGroups,
        ContextUserSummary,
        OuComputers,
        OuGroups,
        OuUsers,
        OuUsersDirectReports,
        OuUsersGroups
    }

    public class ActiveDirectoryQuery
    {
        private static readonly ActiveDirectoryAttribute[]
            DefaultComputerAttributes =
            {
                ActiveDirectoryAttribute.ComputerName,
                ActiveDirectoryAttribute.ComputerDescription,
                ActiveDirectoryAttribute.ComputerDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultComputerGroupsAttributes =
            {
                ActiveDirectoryAttribute.ComputerName,
                ActiveDirectoryAttribute.GroupName,
                ActiveDirectoryAttribute.ComputerDistinguishedName,
                ActiveDirectoryAttribute.GroupDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultGroupAttributes =
            {
                ActiveDirectoryAttribute.GroupName,
                ActiveDirectoryAttribute.GroupManagedBy,
                ActiveDirectoryAttribute.GroupDescription,
                ActiveDirectoryAttribute.GroupDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultGroupComputersAttributes =
            {
                ActiveDirectoryAttribute.GroupName,
                ActiveDirectoryAttribute.ComputerName,
                ActiveDirectoryAttribute.GroupDistinguishedName,
                ActiveDirectoryAttribute.ComputerDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultGroupUsersAttributes =
            {
                ActiveDirectoryAttribute.ContainerGroupName,
                ActiveDirectoryAttribute.UserSurname,
                ActiveDirectoryAttribute.UserGivenName,
                ActiveDirectoryAttribute.UserDisplayName,
                ActiveDirectoryAttribute.UserSamAccountName,
                ActiveDirectoryAttribute.UserIsActive,
                ActiveDirectoryAttribute.UserIsAccountLockedOut,
                ActiveDirectoryAttribute.UserDescription,
                ActiveDirectoryAttribute.UserTitle,
                ActiveDirectoryAttribute.UserCompany,
                ActiveDirectoryAttribute.UserManager,
                ActiveDirectoryAttribute.UserHomeDrive,
                ActiveDirectoryAttribute.UserHomeDirectory,
                ActiveDirectoryAttribute.UserScriptPath,
                ActiveDirectoryAttribute.UserEmailAddress,
                ActiveDirectoryAttribute.UserStreetAddress,
                ActiveDirectoryAttribute.UserCity,
                ActiveDirectoryAttribute.UserState,
                ActiveDirectoryAttribute.UserVoiceTelephoneNumber,
                ActiveDirectoryAttribute.UserPager,
                ActiveDirectoryAttribute.UserMobile,
                ActiveDirectoryAttribute.UserFax,
                ActiveDirectoryAttribute.UserVoip,
                ActiveDirectoryAttribute.UserSip,
                ActiveDirectoryAttribute.UserUserPrincipalName,
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.ContainerGroupDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultGroupUsersDirectReportsAttributes =
            {
                ActiveDirectoryAttribute.ContainerGroupName,
                ActiveDirectoryAttribute.UserName,
                ActiveDirectoryAttribute.DirectReportName,
                ActiveDirectoryAttribute.ContainerGroupDistinguishedName,
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.DirectReportDistinguishedName,
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultUserAttributes =
            {
                ActiveDirectoryAttribute.UserSurname,
                ActiveDirectoryAttribute.UserGivenName,
                ActiveDirectoryAttribute.UserDisplayName,
                ActiveDirectoryAttribute.UserSamAccountName,
                ActiveDirectoryAttribute.UserIsActive,
                ActiveDirectoryAttribute.UserIsAccountLockedOut,
                ActiveDirectoryAttribute.UserDescription,
                ActiveDirectoryAttribute.UserTitle,
                ActiveDirectoryAttribute.UserCompany,
                ActiveDirectoryAttribute.UserManager,
                ActiveDirectoryAttribute.UserHomeDrive,
                ActiveDirectoryAttribute.UserHomeDirectory,
                ActiveDirectoryAttribute.UserScriptPath,
                ActiveDirectoryAttribute.UserEmailAddress,
                ActiveDirectoryAttribute.UserStreetAddress,
                ActiveDirectoryAttribute.UserCity,
                ActiveDirectoryAttribute.UserState,
                ActiveDirectoryAttribute.UserVoiceTelephoneNumber,
                ActiveDirectoryAttribute.UserPager,
                ActiveDirectoryAttribute.UserMobile,
                ActiveDirectoryAttribute.UserFax,
                ActiveDirectoryAttribute.UserVoip,
                ActiveDirectoryAttribute.UserSip,
                ActiveDirectoryAttribute.UserUserPrincipalName,
                ActiveDirectoryAttribute.UserDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultUserDirectReportsAttributes =
            {
                ActiveDirectoryAttribute.UserName,
                ActiveDirectoryAttribute.DirectReportName,
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.DirectReportDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultUserGroupsAttributes =
            {
                ActiveDirectoryAttribute.UserName,
                ActiveDirectoryAttribute.GroupName,
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.GroupDistinguishedName
            };

        private static readonly ActiveDirectoryAttribute[]
            DefaultGroupUsersGroupsAttributes =
            {
                ActiveDirectoryAttribute.ContainerGroupName,
                ActiveDirectoryAttribute.UserName,
                ActiveDirectoryAttribute.GroupName,
                ActiveDirectoryAttribute.ContainerGroupDistinguishedName,
                ActiveDirectoryAttribute.UserDistinguishedName,
                ActiveDirectoryAttribute.GroupDistinguishedName
            };

        private readonly ActiveDirectoryScope _activeDirectoryScope;
        private readonly PrincipalContext _principalContext;
        private readonly string _selectedItemDistinguishedName;

        public ActiveDirectoryQuery(
            QueryType queryType,
            ActiveDirectoryScope activeDirectoryScope = null,
            string selectedItemDistinguishedName = null)
        {
            QueryType = queryType;
            _principalContext = new PrincipalContext(ContextType.Domain);
            _activeDirectoryScope = activeDirectoryScope;
            _selectedItemDistinguishedName = selectedItemDistinguishedName;
        }

        public QueryType QueryType
        {
            get;
        }

        public IEnumerable<ExpandoObject> Data
        {
            get;
            private set;
        }

        private bool QueryTypeIsOu()
        {
            return QueryType == QueryType.OuComputers ||
                   QueryType == QueryType.OuGroups ||
                   QueryType == QueryType.OuUsers ||
                   QueryType == QueryType.OuUsersDirectReports ||
                   QueryType == QueryType.OuUsersGroups;
        }

        private bool QueryTypeIsContextUser()
        {
            return QueryType == QueryType.ContextUserDirectReports ||
                   QueryType == QueryType.ContextUserGroups ||
                   QueryType == QueryType.ContextUserSummary;
        }

        private bool QueryTypeIsContextComputer()
        {
            return QueryType == QueryType.ContextComputerGroups ||
                   QueryType == QueryType.ContextComputerSummary;
        }

        private bool QueryTypeIsContextDirectReport()
        {
            return QueryType == QueryType.ContextDirectReportDirectReports ||
                   QueryType == QueryType.ContextDirectReportGroups ||
                   QueryType == QueryType.ContextDirectReportSummary;
        }

        private bool QueryTypeIsContextGroup()
        {
            return QueryType == QueryType.ContextGroupComputers ||
                   QueryType == QueryType.ContextGroupUsers ||
                   QueryType == QueryType.ContextGroupUsersDirectReports ||
                   QueryType == QueryType.ContextGroupUsersGroups ||
                   QueryType == QueryType.ContextGroupSummary;
        }

        private static IEnumerable<ExpandoObject> GetData(
            DataPreparer dataPreparer)
        {
            return new List<ExpandoObject>(dataPreparer.GetResults());
        }

        private UserPrincipal GetSelectedUserPrincipal()
        {
            return UserPrincipal.FindByIdentity(
                _principalContext, _selectedItemDistinguishedName);
        }

        private GroupPrincipal GetSelectedGroupPrincipal()
        {
            return GroupPrincipal.FindByIdentity(
                _principalContext, _selectedItemDistinguishedName);
        }

        private ComputerPrincipal GetSelectedComputerPrincipal()
        {
            return ComputerPrincipal.FindByIdentity(
                _principalContext, _selectedItemDistinguishedName);
        }

        public async Task Execute()
        {
            DataPreparer dataPreparer = null;
            await Task.Run(() =>
            {
                if (QueryTypeIsOu())
                {
                    var activeDirectorySearcher = new ActiveDirectorySearcher(
                        _activeDirectoryScope);
                    if (QueryType == QueryType.OuComputers)
                    {
                        dataPreparer = new DataPreparer
                        {
                            Data = activeDirectorySearcher.GetComputers(),
                            Attributes = DefaultComputerAttributes
                        };
                    }
                    else if (QueryType == QueryType.OuGroups)
                    {
                        dataPreparer = new DataPreparer
                        {
                            Data = activeDirectorySearcher.GetGroups(),
                            Attributes = DefaultGroupAttributes
                        };
                    }
                    else if (QueryType == QueryType.OuUsers)
                    {
                        dataPreparer = new DataPreparer
                        {
                            Data = activeDirectorySearcher.GetUsers(),
                            Attributes = DefaultUserAttributes
                        };
                    }
                    else if (QueryType == QueryType.OuUsersDirectReports)
                    {
                        dataPreparer = new DataPreparer
                        {
                            Data = activeDirectorySearcher
                                .GetUsersDirectReports(),
                            Attributes = DefaultUserDirectReportsAttributes
                        };
                    }
                    else if (QueryType == QueryType.OuUsersGroups)
                    {
                        dataPreparer = new DataPreparer
                        {
                            Data = activeDirectorySearcher.GetUsersGroups(),
                            Attributes = DefaultUserGroupsAttributes
                        };
                    }
                }
                else if (QueryTypeIsContextComputer())
                {
                    var computerPrincipal = GetSelectedComputerPrincipal();
                    if (QueryType == QueryType.ContextComputerGroups)
                    {
                        dataPreparer = new DataPreparer
                        {
                            Data = new[]
                            {
                                ActiveDirectorySearcher.GetComputerGroups(
                                    computerPrincipal)
                            },
                            Attributes = DefaultComputerGroupsAttributes
                        };
                    }
                }
                else if (QueryTypeIsContextDirectReport() ||
                         QueryTypeIsContextUser())
                {
                    var userPrincipal = GetSelectedUserPrincipal();
                    if (QueryType ==
                        QueryType.ContextDirectReportDirectReports ||
                        QueryType == QueryType.ContextUserDirectReports)
                    {
                        dataPreparer = new DataPreparer
                        {
                            Data = new[]
                            {
                                ActiveDirectorySearcher.GetUserDirectReports(
                                    userPrincipal)
                            },
                            Attributes = DefaultUserDirectReportsAttributes
                        };
                    }
                    else if (QueryType == 
                             QueryType.ContextDirectReportGroups ||
                             QueryType == QueryType.ContextUserGroups)
                    {
                        dataPreparer = new DataPreparer
                        {
                            Data = new[]
                            {
                                ActiveDirectorySearcher.GetUserGroups(
                                    userPrincipal)
                            },
                            Attributes = DefaultUserGroupsAttributes
                        };
                    }
                    else if (QueryType ==
                             QueryType.ContextDirectReportSummary ||
                             QueryType == QueryType.ContextUserSummary)
                    {
                        dataPreparer = new DataPreparer
                        {
                            Data = new[]
                            {
                                userPrincipal
                            },
                            Attributes = DefaultUserAttributes
                        };
                    }
                }
                else if (QueryTypeIsContextGroup())
                {
                    var groupPrincipal = GetSelectedGroupPrincipal();
                    if (QueryType == QueryType.ContextGroupComputers)
                    {
                        dataPreparer = new DataPreparer
                        {
                            Data = new[]
                            {
                                ActiveDirectorySearcher.GetComputers(
                                    groupPrincipal)
                            },
                            Attributes = DefaultGroupComputersAttributes
                        };
                    }
                    else if (QueryType == QueryType.ContextGroupUsers)
                    {
                        dataPreparer = new DataPreparer
                        {
                            Data = new[]
                            {
                                ActiveDirectorySearcher.GetUsers(
                                    groupPrincipal)
                            },
                            Attributes = DefaultGroupUsersAttributes
                        };
                    }
                    else if (QueryType ==
                             QueryType.ContextGroupUsersDirectReports)
                    {
                        dataPreparer = new DataPreparer
                        {
                            Data = new[]
                            {
                                ActiveDirectorySearcher
                                    .GetGroupUsersDirectReports(groupPrincipal)
                            },
                            Attributes =
                                DefaultGroupUsersDirectReportsAttributes
                        };
                    }
                    else if (QueryType == QueryType.ContextGroupUsersGroups)
                    {
                        dataPreparer = new DataPreparer
                        {
                            Data = new[]
                            {
                                ActiveDirectorySearcher.GetGroupUsersGroups(
                                    groupPrincipal)
                            },
                            Attributes = DefaultGroupUsersGroupsAttributes
                        };
                    }
                }
                Data = GetData(dataPreparer);
            });
        }
    }
}