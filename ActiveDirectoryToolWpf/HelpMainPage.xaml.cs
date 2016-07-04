using System.Windows;
using System.Windows.Controls;

namespace ActiveDirectoryToolWpf
{
    /// <summary>
    /// Interaction logic for HelpMainPage.xaml
    /// </summary>
    public partial class HelpMainPage : Page
    {
        public HelpMainPage()
        {
            InitializeComponent();
        }

        private void QueryTypesHyperlink_OnClick(object sender, RoutedEventArgs e)
        {
            NavigationService?.Navigate(new HelpQueryTpesPage());
        }
    }
}
