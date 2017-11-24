using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices.AccountManagement;
using System.Management;
using System.Windows.Threading;
using System.Timers;
using System.Threading;
using System.Diagnostics;
using System.Threading.Tasks;


namespace Join_Domain_Gui_Utility
{
    public partial class JoinDomain : Window
    {
        public JoinDomain()
        {
            InitializeComponent();
            lbl_ProgressBarText.Content = "Please login to Active Directory.";
        }

        //Variables
        //private string computerName { get; set;} not needed
        //private DirectoryEntry existingComputerObject { get; set; }  // did away with using this, cause globals are bad, mmmkay?
        private DirectoryEntry domainRoot { get; set; }
        private String replicationMessage { get; set; }
        //private String 

        private void btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private async Task deletionReplicationThread()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            while (!CheckADReplication("Deletion"))
            {
                if (stopWatch.ElapsedMilliseconds >= 15000)
                {
                    lbl_ProgressBarText.Content = "Checking replication timed out.";
                    MessageBox.Show("Checking replication timed out. Please manually verify that the object has been deleted before joining to the domain.");
                    break;
                }
                Thread.Sleep(2000);
            }

            btn_JoinOrUnjoin.IsEnabled = true;

            stopWatch.Stop();
            //lbl_ProgressBarText.Content = "Checking for complete removal of computer object. ";
            //Stopwatch stopWatch = new Stopwatch();
            //stopWatch.Start();
            //long lastCheck = 0;
            //bool deletionStatus = false;
            //do
            //{
            //    if ((stopWatch.ElapsedMilliseconds - lastCheck) >= 2000)
            //    {
            //        lastCheck = stopWatch.ElapsedMilliseconds;
            //        deletionStatus = CheckADReplication("Deletion");
            //    }
            //    else if (stopWatch.ElapsedMilliseconds >= 15000)
            //    {
            //        lbl_ProgressBarText.Content = "Checking replication timed out.";
            //        MessageBox.Show("Checking replication timed out. Please manually verify that the object has been deleted before joining to the domain.");
            //        break;
            //    }

            //} while (deletionStatus == false);
            //stopWatch.Stop();
            //btn_JoinOrUnjoin.IsEnabled = true;
        }

        private async void creationReplicationThread()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            this.Dispatcher.Invoke(() =>
            {
                while (!CheckADReplication("Creation"))
                {
                    if (stopWatch.ElapsedMilliseconds >= 15000)
                    {
                        lbl_ProgressBarText.Content = "Checking replication timed out.";
                        MessageBox.Show("Checking replication timed out.");
                        break;
                    }
                    Thread.Sleep(2000);
                }
                btn_JoinOrUnjoin.IsEnabled = true;
            });
            stopWatch.Stop();
            //Stopwatch stopWatch = new Stopwatch();
            //stopWatch.Start();
            //long lastCheck = 0;
            //bool creationStatus = false;
            //do
            //{
            //    if ((stopWatch.ElapsedMilliseconds - lastCheck) >= 2000)
            //    {
            //        lastCheck = stopWatch.ElapsedMilliseconds;

            //        creationStatus = CheckADReplication("Creation");
            //    }
            //    else if (stopWatch.ElapsedMilliseconds >= 15000)
            //    {
            //        lbl_ProgressBarText.Content = "Checking replication timed out.";
            //        MessageBox.Show("Checking replication timed out.");
            //        break;
            //    }
            //} while (creationStatus == false);
            //stopWatch.Stop();
            //btn_JoinOrUnjoin.IsEnabled = true;
        }

        private async void btn_JoinOrUnjoin_Click(object sender, RoutedEventArgs e)
        {
            switch (btn_JoinOrUnjoin.Content.ToString())
            {
                case "Delete Object":
                    // Replace with try/catch? nah
                    if (DeleteComputerObject(GetComputerObject(Environment.MachineName)))
                    {
                        btn_JoinOrUnjoin.IsEnabled = false;
                        //Check for replication of object deletion to prevent a mismatched password between the computer object and computer
                        //when joining to the domain immediately after deleting an existing object.
                        lbl_ProgressBarText.Content = "Checking for deletion of computer object.";
                        //Thread t = new Thread(deletionReplicationThread);
                        //t.IsBackground = true;
                        //t.Start();
                        //bool drepl = await deletionReplicationThread();

                        //lbl_ProgressBarText.Content = "Checking for complete removal of computer object. ";
                        //Stopwatch stopWatch = new Stopwatch();
                        //stopWatch.Start();
                        //long lastCheck = 0;
                        //bool deletionStatus = false;
                        //do
                        //{
                        //    if ((stopWatch.ElapsedMilliseconds - lastCheck) >= 2000)
                        //    {
                        //        lastCheck = stopWatch.ElapsedMilliseconds;
                        //        deletionStatus = CheckADReplication("Deletion");
                        //    }
                        //    else if (stopWatch.ElapsedMilliseconds >= 15000)
                        //    {
                        //        lbl_ProgressBarText.Content = "Checking replication timed out.";
                        //        MessageBox.Show("Checking replication timed out. Please manually verify that the object has been deleted before joining to the domain.");
                        //        break;
                        //    }

                        //} while (deletionStatus == false);
                        //while (!CheckADReplication("Creation"))
                        //{
                        //    if (stopWatch.ElapsedMilliseconds >= 15000)
                        //    {
                        //        lbl_ProgressBarText.Content = "Checking replication timed out.";
                        //        MessageBox.Show("Checking replication timed out.");
                        //        break;
                        //    }
                        //    Thread.Sleep(2000);
                        //}
                        //stopWatch.Stop();
                        //btn_JoinOrUnjoin.IsEnabled = true;

                        //If successful, reset GUI elements to allow joining.
                        btn_JoinOrUnjoin.Content = "Join Domain";
                        lbl_ProgressBarText.Content = "Join " + Environment.MachineName + " to ?";
                        tv_OUTree.IsEnabled = true;

                    }
                    else
                    {
                        MessageBox.Show("Could not delete the existing computer object. This is most likely a permissions issue.");
                    }
                    break;
                case "Join Domain":
                    //Check if description and primary operator fields have a value
                    //Also check if the primary operator value is valid, if not, allow user to update the field.
                    //User should not need to click no more than once.
                    MessageBoxResult emptyFieldResult = MessageBoxResult.None;
                    if (tb_Description.Text == string.Empty)
                    {
                        emptyFieldResult = MessageBox.Show("The Description field has been left blank. Do you want to continue joining to the domain?", "Description Empty", MessageBoxButton.YesNo);
                    }
                    else if (tb_PrimaryOperator.Text == string.Empty)
                    {
                        emptyFieldResult = MessageBox.Show("The Primary Operator field has been left blank. Do you want to continue joining to the domain?", "Primary Operator Empty", MessageBoxButton.YesNo);
                    }
                    //If the user did not click no in a previous prompt, continue joining to domain.
                    if (emptyFieldResult != MessageBoxResult.No)
                    {
                        //If primary operator isn't valid, require user to try again, else run code
                        //As long as there is a value for the primary operator, that value needs to be validated.
                        bool isPOValid = false;
                        if (tb_PrimaryOperator.Text != string.Empty)
                        {
                            if (GetPrimaryOperatorDN(tb_PrimaryOperator.Text).ToString() == string.Empty)
                            {
                                MessageBox.Show("The primary operator could not be found. Please check the spelling and try again.");
                                isPOValid = false;
                            }
                            else
                            {
                                isPOValid = true;
                            }
                        }
                        if (tb_PrimaryOperator.Text == string.Empty || isPOValid == true) 
                        {
                            //Make sure to filter out the LDAP:// part of the string!
                            string subString = "OU";
                            string ldapPath = lbl_SelectedOU.Content.ToString();
                            string ouPath = ldapPath.Substring(ldapPath.IndexOf(subString));

                            //Wrap this in an if statement so that the program won't attempt to join if the object wasn't created.
                            bool computerObjectCreated = CreateNewComputerObject(ldapPath);
                            if (computerObjectCreated)
                            {
                                lbl_ProgressBarText.Content = "Successfully created AD object.";
                                lbl_ProgressBarText.Content = "Checking for replication of computer object.";
                                Thread t = new Thread(creationReplicationThread);
                                t.IsBackground = true;
                                t.Start();

                                //Stopwatch stopWatch = new Stopwatch();
                                //stopWatch.Start();
                                //long lastCheck = 0;
                                //bool creationStatus = false;
                                //do
                                //{
                                //    if ((stopWatch.ElapsedMilliseconds - lastCheck) >= 2000)
                                //    {
                                //        lastCheck = stopWatch.ElapsedMilliseconds;

                                //        creationStatus = CheckADReplication("Creation");
                                //    }
                                //    else if (stopWatch.ElapsedMilliseconds >= 15000)
                                //    {
                                //        lbl_ProgressBarText.Content = "Checking replication timed out.";
                                //        MessageBox.Show("Checking replication timed out.");
                                //        break;
                                //    }
                                //} while (creationStatus == false);
                                //stopWatch.Stop();
                                //btn_JoinOrUnjoin.IsEnabled = true;
                                MessageBox.Show("Join domain function temporarily disabled for testing purposes.");
                                lbl_ProgressBarText.Content = "Joining computer to .";

                                //If join occurred successfully, continue running normally, otherwise stop execution and make user try again
                                //int joinResult = JoinDomainMethod(ouPath);
                                int joinResult = 0;
                                if (joinResult == 0)
                                {
                                    //Display a message giving the user the option to restart.
                                    MessageBoxResult mbr = MessageBox.Show(replicationMessage + "\n\nJoining this computer to the domain was successful.\nDo you wish to restart the computer now?", "Restart Options", MessageBoxButton.YesNo);
                                    switch (mbr)
                                    {
                                        case MessageBoxResult.Yes:
                                            System.Diagnostics.Process.Start("shutdown.exe", "/r /t 0");
                                            break;
                                        case MessageBoxResult.No:
                                            Application.Current.Shutdown();
                                            break;
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Could not create the object in active directory. Check your network connection, ensure that the domain servers are reachable, and verify your permissions to create objects in the selected OU.");
                            } //end if else
                        }//end if
                    } //end if
                    break;
            }
        }

        private void btn_Login_Click(object sender, RoutedEventArgs e)
        {
            //Store ad credentials
            string ADUserName = tb_UserName.Text;
            string ADPassword = tb_PasswordBox.Password;

            //Validate AD credentials, if they are valid disable the login button
            if (ValidateADCredentials(ADUserName, ADPassword))
            {
                btn_Login.IsEnabled = false;
                tb_PasswordBox.IsEnabled = false;
                btn_SearchOU.IsEnabled = true;
                tb_SearchOU.IsEnabled = true;
                //btn_ADStatus.IsEnabled = true;
                domainRoot = new DirectoryEntry();
                domainRoot.Path = "LDAP:///ou=,dc=ad,dc=,dc=edu";
                domainRoot.Username = ADUserName;
                domainRoot.Password = ADPassword;
                domainRoot.AuthenticationType = AuthenticationTypes.Secure;

                //Search for existing computer object that matches the current computers name.
                //If existing object is found create delete object gui elements, otherwise initialize the AD Tree
                if (SearchForExistingCN())
                {
                    InitializeTree();
                    LoadDeleteObjectElements();
                    DirectoryEntry de = GetComputerObject(Environment.MachineName);
                    LukeTreeWalker(de.Parent.Path);

                }
                else
                {
                    InitializeTree();
                    lbl_ProgressBarText.Content = "Please select an OU to Join.";
                }
            }
            else
            {
                MessageBox.Show("Invalid username/password. Please re-enter your password and try again.");
            }
        }

        private void InitializeTree()
        {
            DirectorySearcher ouSearcher = new DirectorySearcher(domainRoot);
            ouSearcher.SearchScope = SearchScope.Subtree;
            ouSearcher.PropertiesToLoad.Add("ou");
            ouSearcher.PropertiesToLoad.Add("displayName");
            ouSearcher.PropertiesToLoad.Add("parent");
            ouSearcher.Filter = ("(objectCategory=organizationalUnit)");
            DirectoryEntry rootOU = ouSearcher.FindOne().GetDirectoryEntry();

            tv_OUTree.Items.Add(CreateTreeItem(rootOU));

            //cleanup at the end
            ouSearcher.Dispose();
        }

        private void TreeViewItem_Expanded(object sender, RoutedEventArgs e)
        {
            TreeViewItem item = e.Source as TreeViewItem;
            if ((item.Items.Count == 1) && (item.Items[0] is string))
            {
                item.Items.Clear();
                DirectoryEntry de = new DirectoryEntry(item.Tag.ToString());
                de.Username = tb_UserName.Text;
                de.Password = tb_PasswordBox.Password;
                de.AuthenticationType = AuthenticationTypes.Secure;
                DirectorySearcher ds = new DirectorySearcher(de, "(objectCategory=organizationalUnit)");
                ds.SearchScope = SearchScope.OneLevel;
                foreach (SearchResult sr in ds.FindAll())
                {
                    item.Items.Add(CreateTreeItem(sr.GetDirectoryEntry()));
                }
                item.Items.SortDescriptions.Add(new System.ComponentModel.SortDescription("Header", System.ComponentModel.ListSortDirection.Ascending));
                item.Items.Refresh();
                de.Dispose();
                ds.Dispose();
            }
        }

        private TreeViewItem CreateTreeItem(DirectoryEntry de)
        {
            TreeViewItem item = new TreeViewItem();
            item.Header = de.Properties["ou"][0] + " : ";
            item.Tag = de.Path;
            DirectorySearcher ds = new DirectorySearcher(de, "(objectCategory=organizationalUnit)");
            ds.SearchScope = SearchScope.OneLevel;

            string description = "";
            try
            {
                description = de.Properties["description"][0].ToString();
            }
            catch (Exception e)
            {
            }
            string ouName = "";
            try
            {
                ouName = de.Properties["ou"][0].ToString();
            }
            catch (Exception e)
            {
            }
            item.Header = ouName + "\t\t" + description;
            //If the OU has child items, add a dummy item so that the ou can be expanded.
            if (ds.FindOne() != null)
            {
                item.Items.Add("Loading...");
            }
            de.Dispose();
            ds.Dispose();
            return item;
        }

        private void TreeViewItem_Selected(object sender, RoutedEventArgs e)
        {
            //Enable Join Domain Button
            btn_JoinOrUnjoin.IsEnabled = true;
            //Change progress text
            //lbl_ProgressBarText.Content = "Join to ?"; //Bad place for this.
            TreeViewItem item = e.Source as TreeViewItem;
            //Show path of selected OU in the UI
            string ouIdentifier = "OU";
            string fullPath = item.Tag.ToString();
            lbl_SelectedOU.Content = fullPath.Substring(fullPath.IndexOf(ouIdentifier));

        }

        public bool ValidateADCredentials(string ADUser, string ADPassword)
        {
            PrincipalContext myDomain = new PrincipalContext(ContextType.Domain, "");
            bool isValid = myDomain.ValidateCredentials(ADUser, ADPassword);
            return isValid;
        }

        private bool SearchForExistingCN(DomainController dc, string ADUserName, string ADPassword)
        {

            DirectoryEntry oneDC = new DirectoryEntry();
            oneDC.Path = "LDAP://" + dc.Name + "/ou=,dc=ad,dc=,dc=edu";
            oneDC.Username = ADUserName;
            oneDC.Password = ADPassword;
            oneDC.AuthenticationType = AuthenticationTypes.Secure;
            //oneDC.RefreshCache(); Not needed?
            //Search AD for existing computer object
            String filter = "(&(objectCategory=computer)(name=" + Environment.MachineName + "))";
            //DirectoryEntry de = new DirectoryEntry("LDAP://dc=ad,dc=,dc=edu", ADUserName, ADPassword, AuthenticationTypes.Secure);
            DirectorySearcher ds = new DirectorySearcher(oneDC, filter);
            DirectoryEntry computerObject = null;
            try
            {
                SearchResult sr = ds.FindOne();
                computerObject = sr.GetDirectoryEntry();
            }
            catch (Exception)
            {
                //MessageBox.Show("Search for existing CN -> Something went wrong: " + e.Message);
            }

            ds.Dispose();
            oneDC.Dispose();
            //If computer object already exists return true
            if (computerObject != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private string SearchForOU(string OUName)
        {
            String filter = "(&(objectCategory=organizationalUnit)(name=" + OUName + "))";
            DirectorySearcher ds = new DirectorySearcher(domainRoot, filter);
            string ouPath = "";
            try
            {
                SearchResult sr = ds.FindOne();
                DirectoryEntry de = sr.GetDirectoryEntry();
                ouPath = de.Path.ToString();
                de.Dispose();
            }
            catch (Exception e)
            {
                MessageBox.Show("Could not find the specified OU. Modify your search and try again." + Environment.NewLine + e.Message);
            }
            ds.Dispose();
            return ouPath;
        }

        private DirectoryEntry GetComputerObject(string computerName)
        {
            String filter = "(&(objectCategory=computer)(name=" + computerName + "))";
            DirectorySearcher ds = new DirectorySearcher(domainRoot, filter);
            try
            {
                SearchResult sr = ds.FindOne();
                DirectoryEntry computerObject = sr.GetDirectoryEntry();
                return computerObject;
            }
            catch (Exception e)
            {
                //MessageBox.Show("Search for existing CN -> Something went wrong: " + e.Message);
                return null;
            }
        }

        private bool SearchForExistingCN()
        {
            //Search AD for existing computer object
            String filter = "(&(objectCategory=computer)(name=" + Environment.MachineName + "))";
            DirectorySearcher ds = new DirectorySearcher(domainRoot, filter);
            DirectoryEntry computerObject = new DirectoryEntry();
            try
            {
                SearchResult sr = ds.FindOne();
                computerObject = sr.GetDirectoryEntry();
            }
            catch (Exception)
            {
                computerObject = null;
            }

            ds.Dispose();
            //If computer object already exists return true
            if (computerObject != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private object GetPrimaryOperatorDN(string primaryOperator)
        {
            PrincipalContext myDomain = new PrincipalContext(ContextType.Domain, "", tb_UserName.Text, tb_PasswordBox.Password);
            var sp = new UserPrincipal(myDomain);
            sp.SamAccountName = primaryOperator;

            var userSearch = new PrincipalSearcher();
            userSearch.QueryFilter = sp;

            //Find the user. If user wasn't found display an error message.
            try
            {
                var sr = userSearch.FindOne();
                return sr.DistinguishedName;
            }
            catch (Exception e)
            {
                //MessageBox.Show("Could not find specified user in Active Directory");
                return string.Empty;
            }
        }

        private void LoadDeleteObjectElements()
        {
            btn_JoinOrUnjoin.Content = "Delete Object";
            btn_JoinOrUnjoin.IsEnabled = true;
            tv_OUTree.IsEnabled = false;
            lbl_ProgressBarText.Content = "Computer Object already exists. Delete? Yes/No";
            lbl_ProgressBarText.Visibility = System.Windows.Visibility.Visible;
        }

        private bool DeleteComputerObject(DirectoryEntry computerObject)
        {
            //Delete the existing computer object
            try
            {
                computerObject.DeleteTree();
                computerObject.CommitChanges();
                computerObject.Dispose();
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show("Failed to delete computer object: " + e.Message);
                return false;
            }

        }

        //Not necessary when using the account create flag in the join domain function. 
        //Code to be deprecated after testing but left alone for now in case precreating the object proves to be useful.
        private bool CreateNewComputerObject(string ouPath)
        {
            try
            {
                //Create an object in AD for the computer to bind to.
                //MessageBox.Show(ouPath.ToString());
                //MessageBox.Show(domainRoot.Path.ToString());
                DirectoryEntry ouToJoin = new DirectoryEntry("LDAP:///" + ouPath, tb_UserName.Text, tb_PasswordBox.Password, AuthenticationTypes.Secure);
                DirectoryEntry newComputerObject = ouToJoin.Children.Add("CN=" + Environment.MachineName, "computer");
                newComputerObject.Properties["sAMAccountName"].Value = Environment.MachineName + "$";
                newComputerObject.Properties["UserAccountControl"].Value = 0x1000;
                //The next two fields should only be attempted to be written if the user has provided a value.
                if (tb_Description.Text != "") //|| tb_Description.Text != string.Empty)
                {
                    newComputerObject.Properties["description"].Value = tb_Description.Text;
                }
                if (tb_PrimaryOperator.Text != "")
                {
                    newComputerObject.Properties["managedBy"].Value = GetPrimaryOperatorDN(tb_PrimaryOperator.Text);
                }
                newComputerObject.CommitChanges();

                lbl_ProgressBarText.Content = "Created computer object in AD";
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show("A problem occurred creating the computer object: " + e);
                return false;
            }

        }

        private void WriteDescriptionField(DirectoryEntry computerObject, string description)
        {
            try
            {
                computerObject.Properties["description"].Value = description;
                computerObject.CommitChanges();
                //return true;
            }
            catch (Exception e)
            {
                MessageBox.Show("Could not write computer description field. The following error occurred: \n" + e.Message);
                //return false;
            }
        }

        private void WritePrimaryOperatorField(DirectoryEntry computerObject, object primaryOperator)
        {
            try
            {
                computerObject.Properties["managedBy"].Value = primaryOperator;
                computerObject.CommitChanges();
                //return true;
            }
            catch (Exception e)
            {
                MessageBox.Show("Could not write the managed by field. The following error occurred:\n" + e.Message);
                //return false;
            }
        }

        private int JoinDomainMethod(string ouPath)
        {
            lbl_ProgressBarText.Visibility = System.Windows.Visibility.Visible;
            //Join the computer to the domain
            ManagementObject computerSystem = new ManagementObject("Win32_ComputerSystem.Name='" + Environment.MachineName + "'");
            computerSystem.Scope.Options.Authentication = System.Management.AuthenticationLevel.PacketPrivacy;
            computerSystem.Scope.Options.Impersonation = ImpersonationLevel.Impersonate;
            computerSystem.Scope.Options.EnablePrivileges = true;

            int JOIN_DOMAIN = 1;
            //int ACCT_CREATE = 2;
            //int ACCT_DELETE = 4;
            //int WIN9X_UPGRADE = 16;
            //int DOMAIN_JOIN_IF_JOINED = 32;
            //int JOIN_UNSECURE = 64;
            //int MACHINE_PASSWORD_PASSED = 128;
            //int DEFERRED_SPN_SET = 256;
            //int INSTALL_INVOCATION = 262144;

            string domain = "";
            //string authDomain = ".edu";
            string NetBIOSDN = "";
            string password = tb_PasswordBox.Password;
            string username = tb_UserName.Text;
            string destinationOU = ouPath;
            int parameters = JOIN_DOMAIN;

            object[] methodArgs = { domain, password, NetBIOSDN + "\\" + username, destinationOU, parameters };

            lbl_ProgressBarText.Content = "Binding computer to Active Directory";
            object oResult = computerSystem.InvokeMethod("JoinDomainOrWorkgroup", methodArgs);

            int result = (int)Convert.ToInt32(oResult);

            if (result == 0)
            {
                MessageBox.Show("The computer has been successfully joined to the domain. Please restart to finish the domain join process.");
            }
            else
            {
                string errorDescription = "Error: " + result + ",";
                switch (result)
                {
                    case 5:
                        errorDescription += "Access is denied";
                        break;
                    case 87:
                        errorDescription += "The parameter is incorrect";
                        break;
                    case 110:
                        errorDescription += "The system cannot open the specified object";
                        break;
                    case 1323:
                        errorDescription += "Unable to update the password";
                        break;
                    case 1326:
                        errorDescription += "Logon failure: unknown username or bad password";
                        break;
                    case 1355:
                        errorDescription += "The specified domain either does not exist or could not be contacted";
                        break;
                    case 2224:
                        errorDescription += "The account already exists";
                        break;
                    case 2691:
                        errorDescription += "The machine is already joined to the domain";
                        break;
                    case 2692:
                        errorDescription += "The machine is not currently joined to a domain";
                        break;
                    case 1909:
                        errorDescription += "The referenced account is currently locked out and may not be logged on to";
                        break;
                    default:
                        errorDescription += "Microsoft has not provided a description for this error condition.";
                        break;
                }
                MessageBox.Show(errorDescription);
            }
            return result;
        }

        public List<DomainController> EnumerateDomainControllers()
        {
            DirectoryContext myContext = new DirectoryContext(DirectoryContextType.Domain, "", tb_UserName.Text, tb_PasswordBox.Password);
            List<DomainController> allDCs = new List<DomainController>();
            try
            {
                Domain domain = Domain.GetDomain(myContext);
                foreach (DomainController dc in domain.DomainControllers)
                {
                    allDCs.Add(dc);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred getting the list of domain controllers: " + ex.Message);
            }

            return allDCs;
        }

        //This may not be adequate to verify that all dcs have deleted or added the record, it may not work for both cases.
        private bool CheckADReplication(string mode)
        {
            List<string> messageLines = new List<string>();
            List<bool> replStatus = new List<bool>();
            bool foundCN;
            //List<DomainController> allDCs = EnumerateDomainControllers();
            foreach (DomainController dc in EnumerateDomainControllers())
            {
                //foundCN = SearchForExistingCN(dc, tb_UserName.Text, tb_PasswordBox.Password);
                replStatus.Add(SearchForExistingCN(dc, tb_UserName.Text, tb_PasswordBox.Password));
                messageLines.Add(dc.Name.ToString() + ": " + replStatus.Last().ToString());
            }
            replicationMessage = "Computer object replication status:\n\n";
            replicationMessage += string.Join(Environment.NewLine, messageLines);
            bool replCompletion = false; //Needs initialized so that the return statement is 
            switch (mode)
            {
                case "Deletion":
                    if (replStatus.All(e => e.Equals(false)))
                    {
                        replCompletion = true;
                    }
                    else
                    {
                        replCompletion = false;
                    }

                    break;
                case "Creation":
                    if (replStatus.All(e => e.Equals(true)))
                    {
                        replCompletion = true;
                    }
                    else
                    {
                        replCompletion = false;
                    }
                    break;
                default:
                    break;
            }
            messageLines.Clear(); //Might be necessary to prevent stagnant values/stale data.
            replStatus.Clear(); //Might be necessary to prevent stagnant values/stale data.
            return replCompletion;
        }

        private void CheckADReplication()
        {
            List<string> messageLines = new List<string>();
            List<bool> replStatus = new List<bool>();
            foreach (DomainController dc in EnumerateDomainControllers())
            {
                bool foundCN = SearchForExistingCN(dc, tb_UserName.Text, tb_PasswordBox.Password);
                messageLines.Add(dc.Name.ToString() + ": " + foundCN.ToString());
            }
            replicationMessage = "Computer object replication status:\n\n";
            replicationMessage += string.Join(Environment.NewLine, messageLines);
            MessageBox.Show("Object replication status:" + Environment.NewLine + replicationMessage);
        }


        private bool CheckDeletionReplication()
        {
            List<bool> delStatus = new List<bool>();
            foreach (DomainController dc in EnumerateDomainControllers())
            {
                delStatus.Add(SearchForExistingCN(dc, tb_UserName.Text, tb_PasswordBox.Password));
            }
            if (delStatus.All(e => e.Equals(false)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        //This is not being used. May be implemented in a future version of this program.
        private void btn_ADStatus_Click(object sender, RoutedEventArgs e)
        {
            //ADStatusWindow ADStatus = new ADStatusWindow();
            //ADStatus.DataContext = this;
            //ADStatus.ShowDialog();
            MessageBox.Show("This feature is still in progress.");
        }

        private void btn_SearchOU_Click(object sender, RoutedEventArgs e)
        {
            //If no text in search field, display an error and prompt for input
            //else search for ou and then display the ou in the gui
            if (tb_SearchOU.Text == null || tb_SearchOU.Text == string.Empty)
            {
                MessageBox.Show("No value was provied. Please provide a value then search again.");
            }
            else
            {
                LukeTreeWalker(SearchForOU(tb_SearchOU.Text));
            }
        }

        private void LukeTreeWalker(string OUToSelect)
        {
            if (OUToSelect != string.Empty)
            {
                //Highlight Parent of existing computer object and update lbl_SelectedOU
                string dcIdentifier = ",DC=";
                //string ouIdentifier = "OU=";
                int startIndex = OUToSelect.LastIndexOf("/") + 1;
                int length = OUToSelect.IndexOf(dcIdentifier) - startIndex;
                string ouPath = OUToSelect.Substring(startIndex, length);
                string[] pathArray = ouPath.Split(',');
                TreeViewItem currentItem = new TreeViewItem();
                //Get the single existing treeviewitem from the TreeView gui element and assign it to currentItem
                foreach (TreeViewItem item in tv_OUTree.Items)
                {
                    currentItem = item;
                    //If I don't expand the first item I get a cast problem in the next loop.
                    currentItem.IsExpanded = true;
                }
                int i = 1;
                foreach (string pathItem in pathArray.Reverse())
                {
                    i++;
                    foreach (TreeViewItem tvItem in currentItem.Items)
                    {
                        //MessageBox.Show(item.Substring(item.IndexOf("=") + 1) + " | " + tvItem.Header.ToString());
                        int length2 = tvItem.Tag.ToString().IndexOf(",") - startIndex;
                        string ouName = tvItem.Tag.ToString().Substring(startIndex, length2);
                        if (ouName == pathItem)
                        {
                            currentItem = tvItem;
                            //Don't expand if we're at the last item.
                            //if (i < pathArray.Length)
                            //{
                            //MessageBox.Show(currentItem.Tag.ToString() + Environment.NewLine + pathItem + Environment.NewLine + OUToSelect);
                            currentItem.IsExpanded = true;
                            currentItem.BringIntoView();
                            //}

                        }

                    }
                }
                currentItem.IsSelected = true;
                currentItem.BringIntoView();
                lbl_SelectedOU.Content = OUToSelect.Substring(OUToSelect.LastIndexOf("/") + 1);
            }//end if
        } //end LukeTreeWalker()

    } //end class

} //end main