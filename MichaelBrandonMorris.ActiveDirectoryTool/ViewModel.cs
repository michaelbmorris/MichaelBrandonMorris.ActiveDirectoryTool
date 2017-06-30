using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Deployment.Application;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using GalaSoft.MvvmLight.CommandWpf;
using MichaelBrandonMorris.Extensions.CollectionExtensions;
using MichaelBrandonMorris.Extensions.PrimitiveExtensions;

namespace MichaelBrandonMorris.ActiveDirectoryTool
{
    /// <summary>
    ///     Class ViewModel.
    /// </summary>
    /// <seealso
    ///     cref="System.ComponentModel.INotifyPropertyChanged" />
    /// TODO Edit XML Comment Template for ViewModel
    internal class ViewModel : INotifyPropertyChanged
    {
        /// <summary>
        ///     The computer distinguished name
        /// </summary>
        /// TODO Edit XML Comment Template for ComputerDistinguishedName
        private const string ComputerDistinguishedName =
            "ComputerDistinguishedName";

        /// <summary>
        ///     The container group distinguished name
        /// </summary>
        /// TODO Edit XML Comment Template for ContainerGroupDistinguishedName
        private const string ContainerGroupDistinguishedName =
            "ContainerGroupDistinguishedName";

        /// <summary>
        ///     The container group managed by distinguished name
        /// </summary>
        /// TODO Edit XML Comment Template for ContainerGroupManagedByDistinguishedName
        private const string ContainerGroupManagedByDistinguishedName =
            "ContainerGroupManagedByDistinguishedName";

        /// <summary>
        ///     The direct report distinguished name
        /// </summary>
        /// TODO Edit XML Comment Template for DirectReportDistinguishedName
        private const string DirectReportDistinguishedName =
            "DirectReportDistinguishedName";

        /// <summary>
        ///     The group distinguished name
        /// </summary>
        /// TODO Edit XML Comment Template for GroupDistinguishedName
        private const string GroupDistinguishedName = "GroupDistinguishedName";

        /// <summary>
        ///     The group managed by distinguished name
        /// </summary>
        /// TODO Edit XML Comment Template for GroupManagedByDistinguishedName
        private const string GroupManagedByDistinguishedName =
            "GroupManagedByDistinguishedName";

        /// <summary>
        ///     The help file
        /// </summary>
        /// TODO Edit XML Comment Template for HelpFile
        private const string HelpFile = "ActiveDirectoryToolHelp.chm";

        /// <summary>
        ///     The manager distinguished name
        /// </summary>
        /// TODO Edit XML Comment Template for ManagerDistinguishedName
        private const string ManagerDistinguishedName =
            "ManagerDistinguishedName";

        /// <summary>
        ///     The user distinguished name
        /// </summary>
        /// TODO Edit XML Comment Template for UserDistinguishedName
        private const string UserDistinguishedName = "UserDistinguishedName";

        /// <summary>
        ///     The about window
        /// </summary>
        /// TODO Edit XML Comment Template for _aboutWindow
        private AboutWindow _aboutWindow;

        /// <summary>
        ///     The cancel button visibility
        /// </summary>
        /// TODO Edit XML Comment Template for _cancelButtonVisibility
        private Visibility _cancelButtonVisibility;

        /// <summary>
        ///     The computer search is checked
        /// </summary>
        /// TODO Edit XML Comment Template for _computerSearchIsChecked
        private bool _computerSearchIsChecked;

        /// <summary>
        ///     The context menu items
        /// </summary>
        /// TODO Edit XML Comment Template for _contextMenuItems
        private List<MenuItem> _contextMenuItems;

        /// <summary>
        ///     The context menu visibility
        /// </summary>
        /// TODO Edit XML Comment Template for _contextMenuVisibility
        private Visibility _contextMenuVisibility;

        /// <summary>
        ///     The data
        /// </summary>
        /// TODO Edit XML Comment Template for _data
        private DataView _data;

        /// <summary>
        ///     The group search is checked
        /// </summary>
        /// TODO Edit XML Comment Template for _groupSearchIsChecked
        private bool _groupSearchIsChecked;

        /// <summary>
        ///     The message content
        /// </summary>
        /// TODO Edit XML Comment Template for _messageContent
        private string _messageContent;

        /// <summary>
        ///     The message visibility
        /// </summary>
        /// TODO Edit XML Comment Template for _messageVisibility
        private Visibility _messageVisibility;

        /// <summary>
        ///     The progress bar visibility
        /// </summary>
        /// TODO Edit XML Comment Template for _progressBarVisibility
        private Visibility _progressBarVisibility;

        /// <summary>
        ///     The queries
        /// </summary>
        /// TODO Edit XML Comment Template for _queries
        private Stack<Query> _queries;

        /// <summary>
        ///     The search text
        /// </summary>
        /// TODO Edit XML Comment Template for _searchText
        private string _searchText;

        /// <summary>
        ///     The show distinguished names
        /// </summary>
        /// TODO Edit XML Comment Template for _showDistinguishedNames
        private bool _showDistinguishedNames;

        /// <summary>
        ///     The user search is checked
        /// </summary>
        /// TODO Edit XML Comment Template for _userSearchIsChecked
        private bool _userSearchIsChecked;

        /// <summary>
        ///     The view is enabled
        /// </summary>
        /// TODO Edit XML Comment Template for _viewIsEnabled
        private bool _viewIsEnabled;

        /// <summary>
        ///     Initializes a new instance of the
        ///     <see cref="ViewModel" /> class.
        /// </summary>
        /// TODO Edit XML Comment Template for #ctor
        public ViewModel()
        {
            SetViewVariables();
            try
            {
                RootScope = new ScopeFetcher().Scope;
            }
            catch (PrincipalServerDownException)
            {
                ShowMessage(
                    "The Active Directory server could not be contacted.");
            }

            Queries = new Stack<Query>();
        }

        /// <summary>
        ///     Gets the cancel command.
        /// </summary>
        /// <value>The cancel command.</value>
        /// TODO Edit XML Comment Template for CancelCommand
        public ICommand CancelCommand => new RelayCommand(ExecuteCancel);

        /// <summary>
        ///     Gets the get ou computers command.
        /// </summary>
        /// <value>The get ou computers command.</value>
        /// TODO Edit XML Comment Template for GetOuComputersCommand
        public ICommand GetOuComputersCommand => new RelayCommand(
            ExecuteGetOuComputers,
            CanExecuteOuCommand);

        /// <summary>
        ///     Gets the get ou groups command.
        /// </summary>
        /// <value>The get ou groups command.</value>
        /// TODO Edit XML Comment Template for GetOuGroupsCommand
        public ICommand GetOuGroupsCommand => new RelayCommand(
            ExecuteGetOuGroups,
            CanExecuteOuCommand);

        /// <summary>
        ///     Gets the get ou groups users command.
        /// </summary>
        /// <value>The get ou groups users command.</value>
        /// TODO Edit XML Comment Template for GetOuGroupsUsersCommand
        public ICommand GetOuGroupsUsersCommand => new RelayCommand(
            ExecuteGetOuGroupsUsers,
            CanExecuteOuCommand);

        /// <summary>
        ///     Gets the get ou users command.
        /// </summary>
        /// <value>The get ou users command.</value>
        /// TODO Edit XML Comment Template for GetOuUsersCommand
        public ICommand GetOuUsersCommand => new RelayCommand(
            ExecuteGetOuUsers,
            CanExecuteOuCommand);

        /// <summary>
        ///     Gets the get ou users direct reports command.
        /// </summary>
        /// <value>The get ou users direct reports command.</value>
        /// TODO Edit XML Comment Template for GetOuUsersDirectReportsCommand
        public ICommand GetOuUsersDirectReportsCommand => new RelayCommand(
            ExecuteGetOuUsersDirectReports,
            CanExecuteOuCommand);

        /// <summary>
        ///     Gets the get ou users groups command.
        /// </summary>
        /// <value>The get ou users groups command.</value>
        /// TODO Edit XML Comment Template for GetOuUsersGroupsCommand
        public ICommand GetOuUsersGroupsCommand => new RelayCommand(
            ExecuteGetOuUsersGroups,
            CanExecuteOuCommand);

        /// <summary>
        ///     Gets the open about window.
        /// </summary>
        /// <value>The open about window.</value>
        /// TODO Edit XML Comment Template for OpenAboutWindow
        public ICommand OpenAboutWindow => new RelayCommand(
            ExecuteOpenAboutWindow);

        /// <summary>
        ///     Gets the open help window.
        /// </summary>
        /// <value>The open help window.</value>
        /// TODO Edit XML Comment Template for OpenHelpWindow
        public ICommand OpenHelpWindow => new RelayCommand(
            ExecuteOpenHelpWindow);

        /// <summary>
        ///     Gets the root scope.
        /// </summary>
        /// <value>The root scope.</value>
        /// TODO Edit XML Comment Template for RootScope
        public Scope RootScope
        {
            get;
        }

        /// <summary>
        ///     Gets the run previous query.
        /// </summary>
        /// <value>The run previous query.</value>
        /// TODO Edit XML Comment Template for RunPreviousQuery
        public ICommand RunPreviousQuery => new RelayCommand(
            ExecuteRunPreviousQuery,
            CanExecuteRunPreviousQuery);

        /// <summary>
        ///     Gets the search.
        /// </summary>
        /// <value>The search.</value>
        /// TODO Edit XML Comment Template for Search
        public ICommand Search => new RelayCommand(
            ExecuteSearch,
            CanExecuteSearch);

        /// <summary>
        ///     Gets the search ou.
        /// </summary>
        /// <value>The search ou.</value>
        /// TODO Edit XML Comment Template for SearchOu
        public ICommand SearchOu => new RelayCommand(
            ExecuteSearchOu,
            CanExecuteSearchOu);

        /// <summary>
        ///     Gets the selection changed command.
        /// </summary>
        /// <value>The selection changed command.</value>
        /// TODO Edit XML Comment Template for SelectionChangedCommand
        public ICommand SelectionChangedCommand
        {
            get
            {
                return new RelayCommand<IList>(
                    items =>
                    {
                        SelectedDataRowViews = items;
                    });
            }
        }

        /// <summary>
        ///     Gets the write to file command.
        /// </summary>
        /// <value>The write to file command.</value>
        /// TODO Edit XML Comment Template for WriteToFileCommand
        public ICommand WriteToFileCommand => new RelayCommand(
            ExecuteWriteToFile,
            CanExecuteWriteToFile);

        /// <summary>
        ///     Gets or sets the cancel button visibility.
        /// </summary>
        /// <value>The cancel button visibility.</value>
        /// TODO Edit XML Comment Template for CancelButtonVisibility
        public Visibility CancelButtonVisibility
        {
            get => _cancelButtonVisibility;
            set
            {
                if (_cancelButtonVisibility == value)
                {
                    return;
                }

                _cancelButtonVisibility = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        ///     Gets or sets a value indicating whether [computer
        ///     search is checked].
        /// </summary>
        /// <value>
        ///     <c>true</c> if [computer search is checked];
        ///     otherwise, <c>false</c>.
        /// </value>
        /// TODO Edit XML Comment Template for ComputerSearchIsChecked
        public bool ComputerSearchIsChecked
        {
            get => _computerSearchIsChecked;
            set
            {
                if (_computerSearchIsChecked == value)
                {
                    return;
                }

                _computerSearchIsChecked = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        ///     Gets the context menu items.
        /// </summary>
        /// <value>The context menu items.</value>
        /// TODO Edit XML Comment Template for ContextMenuItems
        public List<MenuItem> ContextMenuItems
        {
            get => _contextMenuItems;
            private set
            {
                if (_contextMenuItems == value)
                {
                    return;
                }

                _contextMenuItems = value;

                ContextMenuVisibility = _contextMenuItems.IsNullOrEmpty()
                    ? Visibility.Hidden
                    : Visibility.Visible;

                NotifyPropertyChanged();
            }
        }

        /// <summary>
        ///     Gets or sets the context menu visibility.
        /// </summary>
        /// <value>The context menu visibility.</value>
        /// TODO Edit XML Comment Template for ContextMenuVisibility
        public Visibility ContextMenuVisibility
        {
            get => _contextMenuVisibility;
            set
            {
                if (_contextMenuVisibility == value)
                {
                    return;
                }

                _contextMenuVisibility = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        ///     Gets or sets the current scope.
        /// </summary>
        /// <value>The current scope.</value>
        /// TODO Edit XML Comment Template for CurrentScope
        public Scope CurrentScope
        {
            get;
            set;
        }

        /// <summary>
        ///     Gets the data.
        /// </summary>
        /// <value>The data.</value>
        /// TODO Edit XML Comment Template for Data
        public DataView Data
        {
            get => _data;
            private set
            {
                if (_data == value)
                {
                    return;
                }

                _data = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        ///     Gets or sets a value indicating whether [group search
        ///     is checked].
        /// </summary>
        /// <value>
        ///     <c>true</c> if [group search is checked]; otherwise,
        ///     <c>false</c>.
        /// </value>
        /// TODO Edit XML Comment Template for GroupSearchIsChecked
        public bool GroupSearchIsChecked
        {
            get => _groupSearchIsChecked;
            set
            {
                if (_groupSearchIsChecked == value)
                {
                    return;
                }

                _groupSearchIsChecked = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        ///     Gets or sets the content of the message.
        /// </summary>
        /// <value>The content of the message.</value>
        /// TODO Edit XML Comment Template for MessageContent
        public string MessageContent
        {
            get => _messageContent;
            set
            {
                if (_messageContent == value)
                {
                    return;
                }

                _messageContent = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        ///     Gets or sets the message visibility.
        /// </summary>
        /// <value>The message visibility.</value>
        /// TODO Edit XML Comment Template for MessageVisibility
        public Visibility MessageVisibility
        {
            get => _messageVisibility;
            set
            {
                if (_messageVisibility == value)
                {
                    return;
                }

                _messageVisibility = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        ///     Gets or sets the progress bar visibility.
        /// </summary>
        /// <value>The progress bar visibility.</value>
        /// TODO Edit XML Comment Template for ProgressBarVisibility
        public Visibility ProgressBarVisibility
        {
            get => _progressBarVisibility;
            set
            {
                if (_progressBarVisibility == value)
                {
                    return;
                }

                _progressBarVisibility = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        ///     Gets or sets the queries.
        /// </summary>
        /// <value>The queries.</value>
        /// TODO Edit XML Comment Template for Queries
        public Stack<Query> Queries
        {
            get => _queries;
            set
            {
                if (_queries == value)
                {
                    return;
                }

                _queries = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        ///     Gets or sets the search text.
        /// </summary>
        /// <value>The search text.</value>
        /// TODO Edit XML Comment Template for SearchText
        public string SearchText
        {
            get => _searchText;
            set
            {
                if (_searchText == value)
                {
                    return;
                }

                _searchText = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        ///     Gets or sets a value indicating whether [show
        ///     distinguished names].
        /// </summary>
        /// <value>
        ///     <c>true</c> if [show distinguished names];
        ///     otherwise, <c>false</c>.
        /// </value>
        /// TODO Edit XML Comment Template for ShowDistinguishedNames
        public bool ShowDistinguishedNames
        {
            get => _showDistinguishedNames;
            set
            {
                if (_showDistinguishedNames == value)
                {
                    return;
                }

                _showDistinguishedNames = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        ///     Gets or sets a value indicating whether [user search is
        ///     checked].
        /// </summary>
        /// <value>
        ///     <c>true</c> if [user search is checked]; otherwise,
        ///     <c>false</c>.
        /// </value>
        /// TODO Edit XML Comment Template for UserSearchIsChecked
        public bool UserSearchIsChecked
        {
            get => _userSearchIsChecked;
            set
            {
                if (_userSearchIsChecked == value)
                {
                    return;
                }

                _userSearchIsChecked = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        ///     Gets the version.
        /// </summary>
        /// <value>The version.</value>
        /// TODO Edit XML Comment Template for Version
        public string Version
        {
            get;
            private set;
        }

        /// <summary>
        ///     Gets a value indicating whether [view is enabled].
        /// </summary>
        /// <value>
        ///     <c>true</c> if [view is enabled]; otherwise,
        ///     <c>false</c>.
        /// </value>
        /// TODO Edit XML Comment Template for ViewIsEnabled
        public bool ViewIsEnabled
        {
            get => _viewIsEnabled;
            private set
            {
                if (_viewIsEnabled == value)
                {
                    return;
                }

                _viewIsEnabled = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        ///     Gets the about window.
        /// </summary>
        /// <value>The about window.</value>
        /// TODO Edit XML Comment Template for AboutWindow
        private AboutWindow AboutWindow
        {
            get
            {
                if (_aboutWindow == null
                    || !_aboutWindow.IsVisible)
                {
                    _aboutWindow = new AboutWindow();
                }

                return _aboutWindow;
            }
        }

        /// <summary>
        ///     Gets the get container groups computers.
        /// </summary>
        /// <value>The get container groups computers.</value>
        /// TODO Edit XML Comment Template for GetContainerGroupsComputers
        private MenuItem GetContainerGroupsComputers => new MenuItem
        {
            Header = "Container Groups - Get Computers",
            Command = new RelayCommand(ExecuteGetContainerGroupsComputers)
        };

        /// <summary>
        ///     Gets the get container groups managed by direct
        ///     reports.
        /// </summary>
        /// <value>The get container groups managed by direct reports.</value>
        /// TODO Edit XML Comment Template for GetContainerGroupsManagedByDirectReports
        private MenuItem GetContainerGroupsManagedByDirectReports => new
            MenuItem
            {
                Header = "Container Group - Get Managed By's Direct Reports",
                Command = new RelayCommand(
                    ExecuteGetContainerGroupsManagedByDirectReports)
            };

        /// <summary>
        ///     Gets the get container groups managed by groups.
        /// </summary>
        /// <value>The get container groups managed by groups.</value>
        /// TODO Edit XML Comment Template for GetContainerGroupsManagedByGroups
        private MenuItem GetContainerGroupsManagedByGroups => new MenuItem
        {
            Header = "Container Groups - Get Managed By's Groups",
            Command = new RelayCommand(ExecuteGetContainerGroupsManagedByGroups)
        };

        /// <summary>
        ///     Gets the get container groups managed by summaries.
        /// </summary>
        /// <value>The get container groups managed by summaries.</value>
        /// TODO Edit XML Comment Template for GetContainerGroupsManagedBySummaries
        private MenuItem GetContainerGroupsManagedBySummaries => new MenuItem
        {
            Header = "Container Groups - Get Managed By's Summaries",
            Command = new RelayCommand(
                ExecuteGetContainerGroupsManagedBySummaries)
        };

        /// <summary>
        ///     Gets the get container groups summaries.
        /// </summary>
        /// <value>The get container groups summaries.</value>
        /// TODO Edit XML Comment Template for GetContainerGroupsSummaries
        private MenuItem GetContainerGroupsSummaries => new MenuItem
        {
            Header = "Container Groups - Get Summaries",
            Command = new RelayCommand(ExecuteGetContainerGroupsSummaries)
        };

        /// <summary>
        ///     Gets the get container groups users.
        /// </summary>
        /// <value>The get container groups users.</value>
        /// TODO Edit XML Comment Template for GetContainerGroupsUsers
        private MenuItem GetContainerGroupsUsers => new MenuItem
        {
            Header = "Container Groups - Get Users",
            Command = new RelayCommand(ExecuteGetContainerGroupsUsers)
        };

        /// <summary>
        ///     Gets the get container groups users direct reports.
        /// </summary>
        /// <value>The get container groups users direct reports.</value>
        /// TODO Edit XML Comment Template for GetContainerGroupsUsersDirectReports
        private MenuItem GetContainerGroupsUsersDirectReports => new MenuItem
        {
            Header = "Container Groups - Get Users' Direct Reports",
            Command = new RelayCommand(
                ExecuteGetContainerGroupsUsersDirectReports)
        };

        /// <summary>
        ///     Gets the get direct reports direct reports.
        /// </summary>
        /// <value>The get direct reports direct reports.</value>
        /// TODO Edit XML Comment Template for GetDirectReportsDirectReports
        private MenuItem GetDirectReportsDirectReports => new MenuItem
        {
            Header = "Direct Reports - Get Direct Reports",
            Command = new RelayCommand(ExecuteGetDirectReportsDirectReports)
        };

        /// <summary>
        ///     Gets the get direct reports groups.
        /// </summary>
        /// <value>The get direct reports groups.</value>
        /// TODO Edit XML Comment Template for GetDirectReportsGroups
        private MenuItem GetDirectReportsGroups => new MenuItem
        {
            Header = "Direct Reports - Get Groups",
            Command = new RelayCommand(ExecuteGetDirectReportsGroups)
        };

        /// <summary>
        ///     Gets the get direct reports summaries.
        /// </summary>
        /// <value>The get direct reports summaries.</value>
        /// TODO Edit XML Comment Template for GetDirectReportsSummaries
        private MenuItem GetDirectReportsSummaries => new MenuItem
        {
            Header = "Direct Reports - Get Summaries",
            Command = new RelayCommand(ExecuteGetDirectReportsSummaries)
        };

        /// <summary>
        ///     Gets the get groups computers.
        /// </summary>
        /// <value>The get groups computers.</value>
        /// TODO Edit XML Comment Template for GetGroupsComputers
        private MenuItem GetGroupsComputers => new MenuItem
        {
            Header = "Groups - Get Computers",
            Command = new RelayCommand(ExecuteGetGroupsComputers)
        };

        /// <summary>
        ///     Gets the get groups summaries.
        /// </summary>
        /// <value>The get groups summaries.</value>
        /// TODO Edit XML Comment Template for GetGroupsSummaries
        private MenuItem GetGroupsSummaries => new MenuItem
        {
            Header = "Groups - Get Summaries",
            Command = new RelayCommand(ExecuteGetGroupsSummaries)
        };

        /// <summary>
        ///     Gets the get groups users.
        /// </summary>
        /// <value>The get groups users.</value>
        /// TODO Edit XML Comment Template for GetGroupsUsers
        private MenuItem GetGroupsUsers => new MenuItem
        {
            Header = "Groups - Get Users",
            Command = new RelayCommand(ExecuteGetGroupsUsers)
        };

        /// <summary>
        ///     Gets the get groups users direct reports.
        /// </summary>
        /// <value>The get groups users direct reports.</value>
        /// TODO Edit XML Comment Template for GetGroupsUsersDirectReports
        private MenuItem GetGroupsUsersDirectReports => new MenuItem
        {
            Header = "Groups - Get Users' Direct Reports",
            Command = new RelayCommand(ExecuteGetGroupsUsersDirectReports)
        };

        /// <summary>
        ///     Gets the get groups users groups.
        /// </summary>
        /// <value>The get groups users groups.</value>
        /// TODO Edit XML Comment Template for GetGroupsUsersGroups
        private MenuItem GetGroupsUsersGroups => new MenuItem
        {
            Header = "Groups - Get Users' Groups",
            Command = new RelayCommand(ExecuteGetGroupsUsersGroups)
        };

        /// <summary>
        ///     Gets the get managers direct reports.
        /// </summary>
        /// <value>The get managers direct reports.</value>
        /// TODO Edit XML Comment Template for GetManagersDirectReports
        private MenuItem GetManagersDirectReports => new MenuItem
        {
            Header = "Managers - Get Direct Reports",
            Command = new RelayCommand(ExecuteGetManagersDirectReports)
        };

        /// <summary>
        ///     Gets the get managers groups.
        /// </summary>
        /// <value>The get managers groups.</value>
        /// TODO Edit XML Comment Template for GetManagersGroups
        private MenuItem GetManagersGroups => new MenuItem
        {
            Header = "Managers - Get Groups",
            Command = new RelayCommand(ExecuteGetManagersGroups)
        };

        /// <summary>
        ///     Gets the get managers summaries.
        /// </summary>
        /// <value>The get managers summaries.</value>
        /// TODO Edit XML Comment Template for GetManagersSummaries
        private MenuItem GetManagersSummaries => new MenuItem
        {
            Header = "Managers - Get Summaries",
            Command = new RelayCommand(ExecuteGetManagersSummaries)
        };

        /// <summary>
        ///     Gets the get users direct reports.
        /// </summary>
        /// <value>The get users direct reports.</value>
        /// TODO Edit XML Comment Template for GetUsersDirectReports
        private MenuItem GetUsersDirectReports => new MenuItem
        {
            Header = "Users - Get Direct Reports",
            Command = new RelayCommand(ExecuteGetUsersDirectReports)
        };

        /// <summary>
        ///     Gets the get users groups.
        /// </summary>
        /// <value>The get users groups.</value>
        /// TODO Edit XML Comment Template for GetUsersGroups
        private MenuItem GetUsersGroups => new MenuItem
        {
            Header = "Users - Get Groups",
            Command = new RelayCommand(ExecuteGetUsersGroups)
        };

        /// <summary>
        ///     Gets the get users summaries.
        /// </summary>
        /// <value>The get users summaries.</value>
        /// TODO Edit XML Comment Template for GetUsersSummaries
        private MenuItem GetUsersSummaries => new MenuItem
        {
            Header = "Users - Get Summaries",
            Command = new RelayCommand(ExecuteGetUsersSummaries)
        };

        /// <summary>
        ///     Gets the menu item get computers groups.
        /// </summary>
        /// <value>The menu item get computers groups.</value>
        /// TODO Edit XML Comment Template for MenuItemGetComputersGroups
        private MenuItem MenuItemGetComputersGroups => new MenuItem
        {
            Header = "Computers - Get Groups",
            Command = new RelayCommand(ExecuteGetComputersGroups)
        };

        /// <summary>
        ///     Gets the menu item get computers summaries.
        /// </summary>
        /// <value>The menu item get computers summaries.</value>
        /// TODO Edit XML Comment Template for MenuItemGetComputersSummaries
        private MenuItem MenuItemGetComputersSummaries => new MenuItem
        {
            Header = "Computers - Get Summaries",
            Command = new RelayCommand(ExecuteGetComputersSummaries)
        };

        /// <summary>
        ///     Gets the menu item get container groups users groups.
        /// </summary>
        /// <value>The menu item get container groups users groups.</value>
        /// TODO Edit XML Comment Template for MenuItemGetContainerGroupsUsersGroups
        private MenuItem MenuItemGetContainerGroupsUsersGroups => new MenuItem
        {
            Header = "Container Groups - Get Users' Groups",
            Command = new RelayCommand(ExecuteGetContainerGroupsUsersGroups)
        };

        /// <summary>
        ///     Gets the menu item get groups managed by direct
        ///     reports.
        /// </summary>
        /// <value>The menu item get groups managed by direct reports.</value>
        /// TODO Edit XML Comment Template for MenuItemGetGroupsManagedByDirectReports
        private MenuItem MenuItemGetGroupsManagedByDirectReports => new
            MenuItem
            {
                Header = "Group - Get Managed By's Direct Reports",
                Command =
                    new RelayCommand(ExecuteGetGroupsManagedByDirectReports)
            };

        /// <summary>
        ///     Gets the menu item get groups managed by groups.
        /// </summary>
        /// <value>The menu item get groups managed by groups.</value>
        /// TODO Edit XML Comment Template for MenuItemGetGroupsManagedByGroups
        private MenuItem MenuItemGetGroupsManagedByGroups => new MenuItem
        {
            Header = "Group - Get Managed By's Groups",
            Command = new RelayCommand(ExecuteGetGroupsManagedByGroups)
        };

        /// <summary>
        ///     Gets the menu item get groups managed by summaries.
        /// </summary>
        /// <value>The menu item get groups managed by summaries.</value>
        /// TODO Edit XML Comment Template for MenuItemGetGroupsManagedBySummaries
        private MenuItem MenuItemGetGroupsManagedBySummaries => new MenuItem
        {
            Header = "Groups - Get Managed By Summaries",
            Command = new RelayCommand(ExecuteGetGroupsManagedBySummaries)
        };

        /// <summary>
        ///     Gets or sets the selected data row views.
        /// </summary>
        /// <value>The selected data row views.</value>
        /// TODO Edit XML Comment Template for SelectedDataRowViews
        private IList SelectedDataRowViews
        {
            get;
            set;
        }

        /// <summary>
        ///     Occurs when a property value changes.
        /// </summary>
        /// TODO Edit XML Comment Template for PropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        ///     Executes the open help window.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteOpenHelpWindow
        private static void ExecuteOpenHelpWindow()
        {
            Process.Start(HelpFile);
        }

        /// <summary>
        ///     Gets the name of the computer distinguished.
        /// </summary>
        /// <param name="dataRowView">The data row view.</param>
        /// <returns>System.String.</returns>
        /// TODO Edit XML Comment Template for GetComputerDistinguishedName
        private static string GetComputerDistinguishedName(
            DataRowView dataRowView)
        {
            return dataRowView[ComputerDistinguishedName].ToString();
        }

        /// <summary>
        ///     Gets the name of the container group distinguished.
        /// </summary>
        /// <param name="dataRowView">The data row view.</param>
        /// <returns>System.String.</returns>
        /// TODO Edit XML Comment Template for GetContainerGroupDistinguishedName
        private static string GetContainerGroupDistinguishedName(
            DataRowView dataRowView)
        {
            return dataRowView[ContainerGroupDistinguishedName].ToString();
        }

        /// <summary>
        ///     Gets the name of the container group managed by
        ///     distinguished.
        /// </summary>
        /// <param name="dataRowView">The data row view.</param>
        /// <returns>System.String.</returns>
        /// TODO Edit XML Comment Template for GetContainerGroupManagedByDistinguishedName
        private static string GetContainerGroupManagedByDistinguishedName(
            DataRowView dataRowView)
        {
            return dataRowView[ContainerGroupManagedByDistinguishedName]
                .ToString();
        }

        /// <summary>
        ///     Gets the name of the direct report distinguished.
        /// </summary>
        /// <param name="dataRowView">The data row view.</param>
        /// <returns>System.String.</returns>
        /// TODO Edit XML Comment Template for GetDirectReportDistinguishedName
        private static string GetDirectReportDistinguishedName(
            DataRowView dataRowView)
        {
            return dataRowView[DirectReportDistinguishedName].ToString();
        }

        /// <summary>
        ///     Gets the name of the group distinguished.
        /// </summary>
        /// <param name="dataRowView">The data row view.</param>
        /// <returns>System.String.</returns>
        /// TODO Edit XML Comment Template for GetGroupDistinguishedName
        private static string GetGroupDistinguishedName(DataRowView dataRowView)
        {
            return dataRowView[GroupDistinguishedName].ToString();
        }

        /// <summary>
        ///     Gets the name of the group managed by distinguished.
        /// </summary>
        /// <param name="dataRowView">The data row view.</param>
        /// <returns>System.String.</returns>
        /// TODO Edit XML Comment Template for GetGroupManagedByDistinguishedName
        private static string GetGroupManagedByDistinguishedName(
            DataRowView dataRowView)
        {
            return dataRowView[GroupManagedByDistinguishedName].ToString();
        }

        /// <summary>
        ///     Gets the name of the manager distinguished.
        /// </summary>
        /// <param name="dataRowView">The data row view.</param>
        /// <returns>System.String.</returns>
        /// TODO Edit XML Comment Template for GetManagerDistinguishedName
        private static string GetManagerDistinguishedName(
            DataRowView dataRowView)
        {
            return dataRowView[ManagerDistinguishedName].ToString();
        }

        /// <summary>
        ///     Gets the name of the user distinguished.
        /// </summary>
        /// <param name="dataRowView">The data row view.</param>
        /// <returns>System.String.</returns>
        /// TODO Edit XML Comment Template for GetUserDistinguishedName
        private static string GetUserDistinguishedName(DataRowView dataRowView)
        {
            return dataRowView[UserDistinguishedName].ToString();
        }

        /// <summary>
        ///     Determines whether this instance [can execute ou
        ///     command].
        /// </summary>
        /// <returns>
        ///     <c>true</c> if this instance [can execute ou
        ///     command]; otherwise, <c>false</c>.
        /// </returns>
        /// TODO Edit XML Comment Template for CanExecuteOuCommand
        private bool CanExecuteOuCommand()
        {
            return CurrentScope != null;
        }

        /// <summary>
        ///     Determines whether this instance [can execute run
        ///     previous query].
        /// </summary>
        /// <returns>
        ///     <c>true</c> if this instance [can execute run previous
        ///     query]; otherwise, <c>false</c>.
        /// </returns>
        /// TODO Edit XML Comment Template for CanExecuteRunPreviousQuery
        private bool CanExecuteRunPreviousQuery()
        {
            return Queries.Multiple();
        }

        /// <summary>
        ///     Determines whether this instance [can execute search].
        /// </summary>
        /// <returns>
        ///     <c>true</c> if this instance [can execute search];
        ///     otherwise, <c>false</c>.
        /// </returns>
        /// TODO Edit XML Comment Template for CanExecuteSearch
        private bool CanExecuteSearch()
        {
            return !SearchText.IsNullOrWhiteSpace() && SearchTypeIsChecked();
        }

        /// <summary>
        ///     Determines whether this instance [can execute search
        ///     ou].
        /// </summary>
        /// <returns>
        ///     <c>true</c> if this instance [can execute search
        ///     ou]; otherwise, <c>false</c>.
        /// </returns>
        /// TODO Edit XML Comment Template for CanExecuteSearchOu
        private bool CanExecuteSearchOu()
        {
            return CanExecuteOuCommand() && CanExecuteSearch();
        }

        /// <summary>
        ///     Determines whether this instance [can execute write to
        ///     file].
        /// </summary>
        /// <returns>
        ///     <c>true</c> if this instance [can execute write to
        ///     file]; otherwise, <c>false</c>.
        /// </returns>
        /// TODO Edit XML Comment Template for CanExecuteWriteToFile
        private bool CanExecuteWriteToFile()
        {
            return Data != null && Data.Count > 0;
        }

        /// <summary>
        ///     Executes the cancel.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteCancel
        private void ExecuteCancel()
        {
            Queries.Peek()?.Cancel();
            ResetQuery();
        }

        /// <summary>
        ///     Executes the get computers groups.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetComputersGroups
        private async void ExecuteGetComputersGroups()
        {
            await RunQuery(
                QueryType.ComputersGroups,
                CurrentScope,
                GetComputersDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get computers summaries.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetComputersSummaries
        private async void ExecuteGetComputersSummaries()
        {
            await RunQuery(
                QueryType.ComputersSummaries,
                CurrentScope,
                GetComputersDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get container groups computers.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetContainerGroupsComputers
        private async void ExecuteGetContainerGroupsComputers()
        {
            await RunQuery(
                QueryType.GroupsComputers,
                CurrentScope,
                GetContainerGroupsDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get container groups managed by direct
        ///     reports.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetContainerGroupsManagedByDirectReports
        private async void ExecuteGetContainerGroupsManagedByDirectReports()
        {
            await RunQuery(
                QueryType.GroupsUsersDirectReports,
                CurrentScope,
                GetContainerGroupsManagedByDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get container groups managed by groups.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetContainerGroupsManagedByGroups
        private async void ExecuteGetContainerGroupsManagedByGroups()
        {
            await RunQuery(
                QueryType.UsersGroups,
                CurrentScope,
                GetContainerGroupsManagedByDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get container groups managed by summaries.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetContainerGroupsManagedBySummaries
        private async void ExecuteGetContainerGroupsManagedBySummaries()
        {
            await RunQuery(
                QueryType.UsersSummaries,
                CurrentScope,
                GetContainerGroupsManagedByDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get container groups summaries.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetContainerGroupsSummaries
        private async void ExecuteGetContainerGroupsSummaries()
        {
            await RunQuery(
                QueryType.GroupsSummaries,
                CurrentScope,
                GetContainerGroupsDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get container groups users.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetContainerGroupsUsers
        private async void ExecuteGetContainerGroupsUsers()
        {
            await RunQuery(
                QueryType.GroupsUsers,
                CurrentScope,
                GetContainerGroupsDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get container groups users direct reports.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetContainerGroupsUsersDirectReports
        private async void ExecuteGetContainerGroupsUsersDirectReports()
        {
            await RunQuery(
                QueryType.GroupsUsersDirectReports,
                CurrentScope,
                GetContainerGroupsDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get container groups users groups.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetContainerGroupsUsersGroups
        private async void ExecuteGetContainerGroupsUsersGroups()
        {
            await RunQuery(
                QueryType.GroupsUsersGroups,
                CurrentScope,
                GetContainerGroupsDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get direct reports direct reports.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetDirectReportsDirectReports
        private async void ExecuteGetDirectReportsDirectReports()
        {
            await RunQuery(
                QueryType.UsersDirectReports,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get direct reports groups.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetDirectReportsGroups
        private async void ExecuteGetDirectReportsGroups()
        {
            await RunQuery(
                QueryType.UsersGroups,
                CurrentScope,
                GetDirectReportsDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get direct reports summaries.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetDirectReportsSummaries
        private async void ExecuteGetDirectReportsSummaries()
        {
            await RunQuery(
                QueryType.UsersSummaries,
                CurrentScope,
                GetDirectReportsDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get groups computers.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetGroupsComputers
        private async void ExecuteGetGroupsComputers()
        {
            await RunQuery(
                QueryType.GroupsComputers,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get groups managed by direct reports.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetGroupsManagedByDirectReports
        private async void ExecuteGetGroupsManagedByDirectReports()
        {
            await RunQuery(
                QueryType.UsersDirectReports,
                CurrentScope,
                GetGroupsManagedByDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get groups managed by groups.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetGroupsManagedByGroups
        private async void ExecuteGetGroupsManagedByGroups()
        {
            await RunQuery(
                QueryType.UsersGroups,
                CurrentScope,
                GetGroupsManagedByDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get groups managed by summaries.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetGroupsManagedBySummaries
        private async void ExecuteGetGroupsManagedBySummaries()
        {
            await RunQuery(
                QueryType.UsersSummaries,
                CurrentScope,
                GetGroupsManagedByDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get groups summaries.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetGroupsSummaries
        private async void ExecuteGetGroupsSummaries()
        {
            await RunQuery(
                QueryType.GroupsSummaries,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get groups users.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetGroupsUsers
        private async void ExecuteGetGroupsUsers()
        {
            await RunQuery(
                QueryType.GroupsUsers,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get groups users direct reports.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetGroupsUsersDirectReports
        private async void ExecuteGetGroupsUsersDirectReports()
        {
            await RunQuery(
                QueryType.GroupsUsersDirectReports,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get groups users groups.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetGroupsUsersGroups
        private async void ExecuteGetGroupsUsersGroups()
        {
            await RunQuery(
                QueryType.GroupsUsersGroups,
                CurrentScope,
                GetGroupsDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get managers direct reports.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetManagersDirectReports
        private async void ExecuteGetManagersDirectReports()
        {
            await RunQuery(
                QueryType.UsersDirectReports,
                CurrentScope,
                GetManagersDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get managers groups.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetManagersGroups
        private async void ExecuteGetManagersGroups()
        {
            await RunQuery(
                QueryType.UsersGroups,
                CurrentScope,
                GetManagersDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get managers summaries.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetManagersSummaries
        private async void ExecuteGetManagersSummaries()
        {
            await RunQuery(
                QueryType.UsersSummaries,
                CurrentScope,
                GetManagersDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get ou computers.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetOuComputers
        private async void ExecuteGetOuComputers()
        {
            await RunQuery(QueryType.OuComputers, CurrentScope);
        }

        /// <summary>
        ///     Executes the get ou groups.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetOuGroups
        private async void ExecuteGetOuGroups()
        {
            await RunQuery(QueryType.OuGroups, CurrentScope);
        }

        /// <summary>
        ///     Executes the get ou groups users.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetOuGroupsUsers
        private async void ExecuteGetOuGroupsUsers()
        {
            await RunQuery(QueryType.OuGroupsUsers, CurrentScope);
        }

        /// <summary>
        ///     Executes the get ou users.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetOuUsers
        private async void ExecuteGetOuUsers()
        {
            await RunQuery(QueryType.OuUsers, CurrentScope);
        }

        /// <summary>
        ///     Executes the get ou users direct reports.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetOuUsersDirectReports
        private async void ExecuteGetOuUsersDirectReports()
        {
            await RunQuery(QueryType.OuUsersDirectReports, CurrentScope);
        }

        /// <summary>
        ///     Executes the get ou users groups.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetOuUsersGroups
        private async void ExecuteGetOuUsersGroups()
        {
            await RunQuery(QueryType.OuUsersGroups, CurrentScope);
        }

        /// <summary>
        ///     Executes the get users direct reports.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetUsersDirectReports
        private async void ExecuteGetUsersDirectReports()
        {
            await RunQuery(
                QueryType.UsersDirectReports,
                CurrentScope,
                GetUsersDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get users groups.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetUsersGroups
        private async void ExecuteGetUsersGroups()
        {
            await RunQuery(
                QueryType.UsersGroups,
                CurrentScope,
                GetUsersDistinguishedNames());
        }

        /// <summary>
        ///     Executes the get users summaries.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteGetUsersSummaries
        private async void ExecuteGetUsersSummaries()
        {
            await RunQuery(
                QueryType.UsersSummaries,
                CurrentScope,
                GetUsersDistinguishedNames());
        }

        /// <summary>
        ///     Executes the open about window.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteOpenAboutWindow
        private void ExecuteOpenAboutWindow()
        {
            AboutWindow.Show();
            AboutWindow.Activate();
        }

        /// <summary>
        ///     Executes the run previous query.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteRunPreviousQuery
        private async void ExecuteRunPreviousQuery()
        {
            Queries.Pop();
            await RunQuery(Queries.Pop());
        }

        /// <summary>
        ///     Executes the search.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteSearch
        private async void ExecuteSearch()
        {
            await RunQuery(GetSearchQueryType(), searchText: SearchText);
        }

        /// <summary>
        ///     Executes the search ou.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteSearchOu
        private async void ExecuteSearchOu()
        {
            await RunQuery(
                GetSearchQueryType(),
                CurrentScope,
                searchText: SearchText);
        }

        /// <summary>
        ///     Executes the write to file.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteWriteToFile
        private async void ExecuteWriteToFile()
        {
            StartTask();
            await Task.Run(
                () =>
                {
                    var fileWriter = new DataFileWriter
                    {
                        Data = Data,
                        Scope = Queries.First().Scope,
                        QueryType = Queries.First().QueryType
                    };

                    ShowMessage("Wrote data to:\n" + fileWriter.WriteToCsv());
                });

            FinishTask();
        }

        /// <summary>
        ///     Finishes the task.
        /// </summary>
        /// TODO Edit XML Comment Template for FinishTask
        private void FinishTask()
        {
            ProgressBarVisibility = Visibility.Hidden;
            CancelButtonVisibility = Visibility.Hidden;
            ViewIsEnabled = true;
        }

        /// <summary>
        ///     Generates the computer context menu items.
        /// </summary>
        /// <returns>IEnumerable&lt;MenuItem&gt;.</returns>
        /// TODO Edit XML Comment Template for GenerateComputerContextMenuItems
        private IEnumerable<MenuItem> GenerateComputerContextMenuItems()
        {
            var computerContextMenuItems = new List<MenuItem>();
            var queryType = Queries.Peek().QueryType;
            if (queryType != QueryType.ComputersGroups)
            {
                computerContextMenuItems.Add(MenuItemGetComputersGroups);
            }

            if (queryType != QueryType.ComputersSummaries)
            {
                computerContextMenuItems.Add(MenuItemGetComputersSummaries);
            }

            return computerContextMenuItems;
        }

        /// <summary>
        ///     Generates the context menu items.
        /// </summary>
        /// <returns>List&lt;MenuItem&gt;.</returns>
        /// TODO Edit XML Comment Template for GenerateContextMenuItems
        private List<MenuItem> GenerateContextMenuItems()
        {
            var contextMenuItems = new List<MenuItem>();
            if (Queries.Peek() == null)
            {
                return null;
            }

            if (Data.Table.Columns.Contains(ComputerDistinguishedName))
            {
                contextMenuItems.AddRange(GenerateComputerContextMenuItems());
            }

            if (Data.Table.Columns.Contains(ContainerGroupDistinguishedName))
            {
                contextMenuItems.Add(GetContainerGroupsComputers);
                contextMenuItems.Add(GetContainerGroupsSummaries);
                contextMenuItems.Add(GetContainerGroupsUsers);
                contextMenuItems.Add(MenuItemGetContainerGroupsUsersGroups);
                contextMenuItems.Add(GetContainerGroupsUsersDirectReports);
            }

            if (Data.Table.Columns.Contains(
                ContainerGroupManagedByDistinguishedName))
            {
                contextMenuItems.Add(GetContainerGroupsManagedByDirectReports);
                contextMenuItems.Add(GetContainerGroupsManagedByGroups);
                contextMenuItems.Add(GetContainerGroupsManagedBySummaries);
            }

            if (Data.Table.Columns.Contains(DirectReportDistinguishedName))
            {
                contextMenuItems.Add(GetDirectReportsDirectReports);
                contextMenuItems.Add(GetDirectReportsGroups);
                contextMenuItems.Add(GetDirectReportsSummaries);
            }

            if (Data.Table.Columns.Contains(GroupDistinguishedName))
            {
                contextMenuItems.AddRange(GenerateGroupContextMenuItems());
            }

            if (Data.Table.Columns.Contains(ManagerDistinguishedName))
            {
                contextMenuItems.Add(GetManagersDirectReports);
                contextMenuItems.Add(GetManagersGroups);
                contextMenuItems.Add(GetManagersSummaries);
            }

            if (!Data.Table.Columns.Contains(UserDistinguishedName))
            {
                return contextMenuItems;
            }

            contextMenuItems.AddRange(GenerateUserContextMenuItems());
            return contextMenuItems;
        }

        /// <summary>
        ///     Generates the group context menu items.
        /// </summary>
        /// <returns>IEnumerable&lt;MenuItem&gt;.</returns>
        /// TODO Edit XML Comment Template for GenerateGroupContextMenuItems
        private IEnumerable<MenuItem> GenerateGroupContextMenuItems()
        {
            var groupContextMenuItems = new List<MenuItem>();
            var queryType = Queries.Peek().QueryType;
            if (queryType != QueryType.GroupsComputers)
            {
                groupContextMenuItems.Add(GetGroupsComputers);
            }

            if (queryType != QueryType.GroupsSummaries)
            {
                groupContextMenuItems.Add(GetGroupsSummaries);
            }

            if (queryType != QueryType.GroupsUsers)
            {
                groupContextMenuItems.Add(GetGroupsUsers);
            }

            if (queryType != QueryType.GroupsUsersDirectReports)
            {
                groupContextMenuItems.Add(GetGroupsUsersDirectReports);
            }

            if (queryType != QueryType.GroupsUsersGroups)
            {
                groupContextMenuItems.Add(GetGroupsUsersGroups);
            }

            if (Data.Table.Columns.Contains(GroupManagedByDistinguishedName))
            {
                groupContextMenuItems.AddRange(
                    GenerateGroupManagedByContextMenuItems());
            }

            return groupContextMenuItems;
        }

        /// <summary>
        ///     Generates the group managed by context menu items.
        /// </summary>
        /// <returns>IEnumerable&lt;MenuItem&gt;.</returns>
        /// TODO Edit XML Comment Template for GenerateGroupManagedByContextMenuItems
        private IEnumerable<MenuItem> GenerateGroupManagedByContextMenuItems()
        {
            return new List<MenuItem>
            {
                MenuItemGetGroupsManagedByDirectReports,
                MenuItemGetGroupsManagedByGroups,
                MenuItemGetGroupsManagedBySummaries
            };
        }

        /// <summary>
        ///     Generates the user context menu items.
        /// </summary>
        /// <returns>IEnumerable&lt;MenuItem&gt;.</returns>
        /// TODO Edit XML Comment Template for GenerateUserContextMenuItems
        private IEnumerable<MenuItem> GenerateUserContextMenuItems()
        {
            var userContextMenuItems = new List<MenuItem>();
            var queryType = Queries.Peek().QueryType;
            if (queryType != QueryType.UsersDirectReports)
            {
                userContextMenuItems.Add(GetUsersDirectReports);
            }

            if (queryType != QueryType.UsersGroups)
            {
                userContextMenuItems.Add(GetUsersGroups);
            }

            if (queryType != QueryType.UsersSummaries)
            {
                userContextMenuItems.Add(GetUsersSummaries);
            }

            return userContextMenuItems;
        }

        /// <summary>
        ///     Gets the computers distinguished names.
        /// </summary>
        /// <returns>IEnumerable&lt;System.String&gt;.</returns>
        /// TODO Edit XML Comment Template for GetComputersDistinguishedNames
        private IEnumerable<string> GetComputersDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>()
                .Select(GetComputerDistinguishedName)
                .Where(c => !c.IsNullOrWhiteSpace());
        }

        /// <summary>
        ///     Gets the container groups distinguished names.
        /// </summary>
        /// <returns>IEnumerable&lt;System.String&gt;.</returns>
        /// TODO Edit XML Comment Template for GetContainerGroupsDistinguishedNames
        private IEnumerable<string> GetContainerGroupsDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>()
                .Select(GetContainerGroupDistinguishedName)
                .Where(c => !c.IsNullOrWhiteSpace());
        }

        /// <summary>
        ///     Gets the container groups managed by distinguished
        ///     names.
        /// </summary>
        /// <returns>IEnumerable&lt;System.String&gt;.</returns>
        /// TODO Edit XML Comment Template for GetContainerGroupsManagedByDistinguishedNames
        private IEnumerable<string>
            GetContainerGroupsManagedByDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>()
                .Select(GetContainerGroupManagedByDistinguishedName)
                .Where(s => !s.IsNullOrWhiteSpace());
        }

        /// <summary>
        ///     Gets the direct reports distinguished names.
        /// </summary>
        /// <returns>IEnumerable&lt;System.String&gt;.</returns>
        /// TODO Edit XML Comment Template for GetDirectReportsDistinguishedNames
        private IEnumerable<string> GetDirectReportsDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>()
                .Select(GetDirectReportDistinguishedName)
                .Where(d => !d.IsNullOrWhiteSpace());
        }

        /// <summary>
        ///     Gets the groups distinguished names.
        /// </summary>
        /// <returns>IEnumerable&lt;System.String&gt;.</returns>
        /// TODO Edit XML Comment Template for GetGroupsDistinguishedNames
        private IEnumerable<string> GetGroupsDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>()
                .Select(GetGroupDistinguishedName)
                .Where(g => !g.IsNullOrWhiteSpace());
        }

        /// <summary>
        ///     Gets the groups managed by distinguished names.
        /// </summary>
        /// <returns>IEnumerable&lt;System.String&gt;.</returns>
        /// TODO Edit XML Comment Template for GetGroupsManagedByDistinguishedNames
        private IEnumerable<string> GetGroupsManagedByDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>()
                .Select(GetGroupManagedByDistinguishedName)
                .Where(s => !s.IsNullOrWhiteSpace());
        }

        /// <summary>
        ///     Gets the managers distinguished names.
        /// </summary>
        /// <returns>IEnumerable&lt;System.String&gt;.</returns>
        /// TODO Edit XML Comment Template for GetManagersDistinguishedNames
        private IEnumerable<string> GetManagersDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>()
                .Select(GetManagerDistinguishedName)
                .Where(m => !m.IsNullOrWhiteSpace());
        }

        /// <summary>
        ///     Gets the type of the search query.
        /// </summary>
        /// <returns>QueryType.</returns>
        /// TODO Edit XML Comment Template for GetSearchQueryType
        private QueryType GetSearchQueryType()
        {
            if (ComputerSearchIsChecked)
            {
                return QueryType.SearchComputer;
            }

            if (GroupSearchIsChecked)
            {
                return QueryType.SearchGroup;
            }

            return UserSearchIsChecked ? QueryType.SearchUser : QueryType.None;
        }

        /// <summary>
        ///     Gets the users distinguished names.
        /// </summary>
        /// <returns>IEnumerable&lt;System.String&gt;.</returns>
        /// TODO Edit XML Comment Template for GetUsersDistinguishedNames
        private IEnumerable<string> GetUsersDistinguishedNames()
        {
            return SelectedDataRowViews.Cast<DataRowView>()
                .Select(GetUserDistinguishedName)
                .Where(u => !u.IsNullOrWhiteSpace());
        }

        /// <summary>
        ///     Hides the message.
        /// </summary>
        /// TODO Edit XML Comment Template for HideMessage
        private void HideMessage()
        {
            MessageContent = string.Empty;
            MessageVisibility = Visibility.Hidden;
        }

        /// <summary>
        ///     Notifies the property changed.
        /// </summary>
        /// <param name="propertyName">Name of the property.</param>
        /// TODO Edit XML Comment Template for NotifyPropertyChanged
        private void NotifyPropertyChanged(
            [CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(
                this,
                new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        ///     Resets the query.
        /// </summary>
        /// TODO Edit XML Comment Template for ResetQuery
        private void ResetQuery()
        {
            if (Queries.Any())
            {
                Queries.Pop();
            }
        }

        /// <summary>
        ///     Runs the query.
        /// </summary>
        /// <param name="queryType">Type of the query.</param>
        /// <param name="scope">The scope.</param>
        /// <param name="selectedItemDistinguishedNames">
        ///     The selected
        ///     item distinguished names.
        /// </param>
        /// <param name="searchText">The search text.</param>
        /// <returns>Task.</returns>
        /// TODO Edit XML Comment Template for RunQuery
        private async Task RunQuery(
            QueryType queryType,
            Scope scope = null,
            IEnumerable<string> selectedItemDistinguishedNames = null,
            string searchText = null)
        {
            await RunQuery(
                new Query(
                    queryType,
                    scope,
                    selectedItemDistinguishedNames?.ToList(),
                    searchText));
        }

        /// <summary>
        ///     Runs the query.
        /// </summary>
        /// <param name="query">The query.</param>
        /// <returns>Task.</returns>
        /// TODO Edit XML Comment Template for RunQuery
        private async Task RunQuery(Query query)
        {
            StartTask();

            try
            {
                Queries.Push(query);
                await Queries.Peek().Execute();
                Data = Queries.Peek().Data.ToDataTable().AsDataView();
                ContextMenuItems = GenerateContextMenuItems();
            }
            catch (NullReferenceException e)
            {
                ShowMessage(e.Message);
                ResetQuery();
            }
            catch (OperationCanceledException)
            {
                ShowMessage("Operation was cancelled.");
                ResetQuery();
            }
            catch (ArgumentNullException)
            {
                ShowMessage(
                    "No results of desired type found in selected context.");

                ResetQuery();
            }
            catch (OutOfMemoryException)
            {
                ShowMessage("The selected query is too large to run.");
                ResetQuery();
            }

            FinishTask();
        }

        /// <summary>
        ///     Searches the type is checked.
        /// </summary>
        /// <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
        /// TODO Edit XML Comment Template for SearchTypeIsChecked
        private bool SearchTypeIsChecked()
        {
            return ComputerSearchIsChecked
                   || GroupSearchIsChecked
                   || UserSearchIsChecked;
        }

        /// <summary>
        ///     Sets the view variables.
        /// </summary>
        /// TODO Edit XML Comment Template for SetViewVariables
        private void SetViewVariables()
        {
            ProgressBarVisibility = Visibility.Hidden;
            MessageVisibility = Visibility.Hidden;
            CancelButtonVisibility = Visibility.Hidden;
            ViewIsEnabled = true;
            ContextMenuItems = null;

            try
            {
                Version = ApplicationDeployment.CurrentDeployment.CurrentVersion
                    .ToString();
            }
            catch (InvalidDeploymentException)
            {
                Version = "Not Installed";
            }
        }

        /// <summary>
        ///     Shows the message.
        /// </summary>
        /// <param name="messageContent">Content of the message.</param>
        /// TODO Edit XML Comment Template for ShowMessage
        private void ShowMessage(string messageContent)
        {
            MessageContent = messageContent + "\n\nDouble-click to dismiss.";
            MessageVisibility = Visibility.Visible;
        }

        /// <summary>
        ///     Starts the task.
        /// </summary>
        /// TODO Edit XML Comment Template for StartTask
        private void StartTask()
        {
            HideMessage();
            ViewIsEnabled = false;
            ProgressBarVisibility = Visibility.Visible;
            CancelButtonVisibility = Visibility.Visible;
        }
    }
}