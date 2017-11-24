using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Security.Principal;
using System.DirectoryServices.ActiveDirectory;
using System.Management;

namespace Join_Domain_Gui_Utility
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            //Check that the program has been run with admin credentials in the event of UAC being disabled.
            if (!IsUserAdministrator())
            {
                MessageBox.Show("Please run this program as an administrator!");
                Application.Current.Shutdown();
            }
            //Check for network connectivity before continuing
            bool networkUp;
            do
            {
                networkUp = System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable();
                if (!networkUp)
                {
                    MessageBoxResult networkCheckPrompt = MessageBox.Show("An active network connection could not be detected. Click \'OK\' to continue after establishing a network connection or click \'Cancel\' to quit.", "No Active Connection", MessageBoxButton.OKCancel);
                    if (networkCheckPrompt == MessageBoxResult.Cancel)
                    {
                        Application.Current.Shutdown();
                    }
                }
            } while (!networkUp);

            if (IsDomainJoined())
            {
                MessageBoxResult domainJoinedPrompt = MessageBox.Show("This computer is already joined to a domain. Please unjoin the computer and then run this utility again.", "Computer Already Joined", MessageBoxButton.OK);
                Application.Current.Shutdown();
            }

            lbl_ComputerName.Content = Environment.MachineName;
        }

        // Variables


        private void btn_No_Click(object sender, RoutedEventArgs e)
        {
            ClickedNo ClickedNo = new ClickedNo();
            ClickedNo.ShowDialog();
        }

        private void btn_Yes_Click(object sender, RoutedEventArgs e)
        {
            JoinDomain JoinWindow = new JoinDomain();
            JoinWindow.Show();
            this.Close();
        }

        public bool IsUserAdministrator()
        {
            bool isAdmin;
            try
            {
                WindowsIdentity user = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(user);
                isAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch (UnauthorizedAccessException ex)
            {
                isAdmin = false;
            }
            catch (Exception ex)
            {
                isAdmin = false;
            }
            return isAdmin;
        }

        public bool IsDomainJoined()
        {
            ManagementObject ComputerSystem;
            using (ComputerSystem = new ManagementObject(String.Format("Win32_ComputerSystem.Name='{0}'", Environment.MachineName)))
            {
                ComputerSystem.Get();
                object Result = ComputerSystem["PartOfDomain"];
                return (Result != null && (bool)Result);
            }
        }

    }
}
