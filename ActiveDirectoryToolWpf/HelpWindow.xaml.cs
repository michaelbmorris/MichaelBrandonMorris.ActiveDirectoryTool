namespace ActiveDirectoryToolWpf
{
    public partial class HelpWindow
    {
        public HelpWindow()
        {
            InitializeComponent();
            NavigationFrame.Navigate(new HelpMainPage());
        }
    }
}
