using System;
using System.Collections.Generic;
using System.Management;
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
using Microsoft.Win32;
using System.Net;
using System.Net.NetworkInformation;
using System.Security.Principal;
using System.DirectoryServices.AccountManagement;
using System.Diagnostics;
using System.Xml;

namespace build_certificate
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static string ComputerName()
        {
            string result = string.Empty;
            {
                try
                {
                    if (result != null)
                        result = Environment.MachineName;
                    if (result == null)
                        result = "Unable to obtain";
                }
                catch (Exception ex)
                {
                    result = "Unable to obtain";
                }
                return result;
            }
        }
        //Calls WMI and searches for OS caption name
        public static string GetOSFriendlyName()
        {
            string result = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Caption FROM Win32_OperatingSystem");
                foreach (ManagementObject os in searcher.Get())
                {
                    result = os["Caption"].ToString();
                    break;
                }
            }
            catch (Exception ex)
            {
                result = "Unable to obtain";
            }
            return result;
        }
        //Calls WMI and searches for SMBIOSAssetTag
        public static string GetAssetTag()
        {
            string result = string.Empty;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_SystemEnclosure");
            foreach (ManagementObject SMBIOSAssetTag in searcher.Get())
                try
                {
                    result = SMBIOSAssetTag["SMBIOSAssetTag"].ToString();
                    if (result == null || result == "")
                        result = "Unable to obtain";
                }
                catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                {
                    //react appropriately
                    result = "Unable to obtain";
                    break;
                }
            return result;
        }
        //Calls WMI and searches for InstallDate
        public static string InstallDate()
        {
            string result = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                foreach (ManagementObject InstallDate in searcher.Get())
                {
                    DateTime dt = ManagementDateTimeConverter.ToDateTime(InstallDate["InstallDate"].ToString());
                    result = dt.ToString();
                    break;
                }
            }
            catch (Exception ex)
            {
                result = "Unable to obtain";
            }
            return result;
        }
        //Calls WMI and searches for Build Version (Release ID and build number)
        public static string BuildNumber()
        {
            string build = string.Empty;
            string result = string.Empty;
            string version = string.Empty;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
            foreach (ManagementObject BuildNumber in searcher.Get())
            {
                build = BuildNumber["BuildNumber"].ToString();
                build = build.Insert(0, "Build "); //Insert "Build" at the first index
                try //now we check the registry key values for the Release ID
                {
                    using (RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Windows NT\\CurrentVersion"))
                    {
                        if (key != null)
                        {
                            Object o = key.GetValue("ReleaseID");
                            if (o != null)
                            {
                                version = o.ToString(); //convert object to string
                            }
                        }
                    }
                }
                catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                {
                    //react appropriately
                }
                version = version.Insert(0, "Version "); //Insert "Version" at the first index
                result = version + " " + build;
                break;
            }
            return result;
        }
        //Calls WMI and searches for Kernel Version
        public static string Kernel()
        {
            string result = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                foreach (ManagementObject Kernel in searcher.Get())
                {
                    result = Kernel["Version"].ToString();
                    break;
                }
            }
            catch (Exception ex)
            {
                result = "Unable to obtain";
            }
            return result;
        }
        //Calls WMI and searches for IE Version
        public static string IEVersion()
        {
            string result = string.Empty;
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Internet Explorer"))
                {
                    if (key != null)
                    {
                        Object o = key.GetValue("svcVersion");
                        if (o != null)
                        {
                            result = o.ToString(); //convert object to string
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                result = ex.ToString();
            }
            return result;
        }

        //Calls Registry Key and searches for Microsoft Office Version
        public static string MSOfficeVersion()
        {
            string result = string.Empty;
            //var fileVersionInfo = "";
            //var version = "";
                try
                {
                var fileVersionInfo = FileVersionInfo.GetVersionInfo(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");
                var version = new Version(fileVersionInfo.FileVersion);
                result = version.ToString();
                if (result == null || result == "")
                        result = "Unable to obtain";
                }
                catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                {
                    //react appropriately
                    result = "Unable to obtain";
                }
            return result;
        }

        ////Calls Registry Key and searches for Microsoft Office Version
        //public static string MSOfficeVersion()
        //{
        //    string result = string.Empty;
        //    ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Version FROM Win32_Product WHERE Name = 'Office 16 Click-to-Run Licensing Component'");
        //    foreach (ManagementObject Version in searcher.Get())
        //        try
        //        {
        //            result = Version["Version"].ToString();
        //            if (result == null || result == "")
        //                result = "Unable to obtain";
        //        }
        //        catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
        //        {
        //            //react appropriately
        //            result = "Unable to obtain";
        //            break;
        //        }
        //    return result;
        //}

        //Calls WMI and searches for AppV Version
        public static string AppVVersion()
        {
            string result = string.Empty;
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\AppV\\Client"))
                {
                    if (key != null)
                    {
                        Object o = key.GetValue("Version");
                        if (o != null)
                        {
                            result = o.ToString(); //convert object to string
                        }
                    }
                    if (key == null)
                    {
                        result = "No AppV Detected";
                    }
                }
            }
            catch (Exception ex)
            {
                result = ex.ToString();
            }
            return result;
        }
        //
        //!--Hardware Info--!!
        //
        //Calls WMI and searches for Manufacturer
        public static string MakeName()
        {
            string result = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_BIOS");
                foreach (ManagementObject MakeName in searcher.Get())
                {
                    result = MakeName["Manufacturer"].ToString();
                    break;
                }
            }
            catch (Exception ex)
            {
                result = "Unable to obtain";
            }

            return result;
        }
        //Calls WMI and searches for Model
        public static string ModelName()
        {
            string result = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
                foreach (ManagementObject ModelName in searcher.Get())
                {
                    result = ModelName["Model"].ToString();
                    break;
                }
            }
            catch (Exception ex)
            {
                result = "Unable to obtain";
            }
            return result;
        }
        //Calls WMI and searches for BIOSVersion
        public static string UEFIBIOSVersion()
        {
            string result = string.Empty;

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Bios");
            foreach (ManagementObject UEFIBIOSVersion in searcher.Get())
            {
                result = UEFIBIOSVersion["SMBIOSBIOSVersion"].ToString();
                break;
            }
            return result;
        }
        //Calls WMI and searches for SystemSKU
        public static string SKUNumber()
        {
            string result = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
                foreach (ManagementObject SKUNumber in searcher.Get())
                {
                    result = SKUNumber["SystemSkuNumber"].ToString();
                    break;
                }
            }
            catch (Exception ex)
            {
                result = "Unable to obtain";
            }
            return result;
        }
        //Calls WMI and searches for Name
        public static string ProcessorName()
        {
            string result = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
                foreach (ManagementObject ProcessorName in searcher.Get())
                {
                    result = ProcessorName["Name"].ToString();
                    break;
                }
            }
            catch (Exception ex)
            {
                result = "Unable to obtain";
            }
            return result;
        }
        //Calls WMI and searches for TotalVisibleMemorySize
        public static string PhysicalMemory()
        {
            double memval;
            double result = 0;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                foreach (ManagementObject PhysicalMemory in searcher.Get())
                {
                    memval = Convert.ToDouble(PhysicalMemory["TotalVisibleMemorySize"]);
                    result = Math.Round((memval / (1024 * 1024)), 0);
                    break;
                }
            }
            catch (Exception ex)
            {
                result.ToString("Unable to obtain");
            }
            return result.ToString(result+" GB");
        }
        //Calls DNS class to return IP for hostname in Index 1 (Index 0 is normally loopback)
        public static string IPAddress()
        {
            string hostName = Dns.GetHostName();
            string result = Dns.GetHostEntry(hostName).AddressList[0].ToString();

            return result;
        }
        //Calls Physical Address Class to return MAC Address
                public static string MACAddress()
        {
            string result = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE INDEX=1");
                foreach (ManagementObject MACAddress in searcher.Get())
                {
                    result = MACAddress["MACAddress"].ToString();
                    if (result == null || result == "")
                        result = "Unable to obtain";

                    break;
                }
            }
            catch (Exception ex)
            {
                result = "Unable to obtain";
            }
            return result;
        }

        //Calls WMI and searches for GatewayAddress
        public static string GatewayAddress()
        {
            string result = string.Empty;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT DefaultIPGateway FROM Win32_NetworkAdapterConfiguration WHERE INDEX=1");
            foreach (ManagementObject GatewayAddress in searcher.Get())
            {
                try
                {
                    //result = GatewayAddress["DefaultIPGateway"].GetType().ToString();
                    Object obj = GatewayAddress["DefaultIPGateway"];                                                    //my code works and I dont' know why...
                    string[] array = (obj as IEnumerable<object>).Cast<object>().Select(x => x.ToString()).ToArray();
                    result = array[0];
                    //result = obj[0].ToString;                                                                         //my code doesn't work and I don't know why...
                }
                catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                {
                    //react appropriately
                    if (result == null || result == "")
                        result = "Unable to obtain";
                        break;
                }

                break;
            }
            return result;
        }
        //Calls NetworkInterface class and looks for first network card and grabs it's DNS address
                public static string DnsAddress()
        {
            string result = string.Empty;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT DNSServerSearchOrder FROM Win32_NetworkAdapterConfiguration WHERE INDEX=1");
            foreach (ManagementObject DnsAddress in searcher.Get())
            {
                try
                {
                    //result = GatewayAddress["DefaultIPGateway"].GetType().ToString();
                    Object obj = DnsAddress["DNSServerSearchOrder"];                                                    //my code works and I dont' know why...
                    string[] array = (obj as IEnumerable<object>).Cast<object>().Select(x => x.ToString()).ToArray();
                    result = array[0];
                    //result = obj[0].ToString;                                                                         //my code doesn't work and I don't know why...
                }
                catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                {
                    //react appropriately
                    if (result == null || result == "")
                        result = "Unable to obtain";
                        break;
                }

                break;
            }
            return result;
        }

        //
        //!--Domain Info--!!
        //

        //Calls WMI and searches for DomainName
        public static string DomainName()
        {
            string result = string.Empty;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
            foreach (ManagementObject DomainName in searcher.Get())
            {
                try
                {
                    result = DomainName["Domain"].ToString();
                }
                catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                {
                    //react appropriately
                    if (result == null)
                        result = "Unable to obtain";
                        break;
                }

                break;
            }
            return result;
        }
        //Calls WMI and searches for UserName
        public static string UserName()
        {
            string result = string.Empty;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
            foreach (ManagementObject Username in searcher.Get())
            {
                try
                {
                    result = Username["UserName"].ToString();
                }
                catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                {
                    //react appropriately
                    if (result == null)
                        result = "Unable to obtain";
                        break;
                }

                break;
            }
            return result;
        }
        //Calls WMI and searches for ClientSiteName
        public static string ADSiteName()
        {
            string result = string.Empty;
            {
                try
                {
                    result = (System.DirectoryServices.ActiveDirectory.ActiveDirectorySite.GetComputerSite().ToString());
                    if (result == null || result == "")
                    result = "Unable to obtain";
                }
                catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                {
                    //react appropriately
                    if (result == null || result == "")
                        result = "Unable to obtain";
                }

            }
            return result;
        }

        //Calls environment variable and searches for logonserver
        public static string DCServerName()
        {
            string result = string.Empty;
            {
                try
                {
                    result = Environment.GetEnvironmentVariable("logonserver");
                    if (result == null)
                        result = "Unable to obtain";
                }
                catch (Exception ex)
                {
                    result = "Unable to obtain";
                }
                return result;
            }
        }
        //Calls environment variable SubAreaID and returns it
        public static string SubAreaName()
        {
            string result = string.Empty;
            {
                try
                {
                    if (result != null)
                        result = Environment.GetEnvironmentVariable("SubAreaID");
                    if (result == null)
                        result = "Unable to obtain";
                }
                catch (Exception ex)
                {
                    result = "Unable to obtain";
                }
                return result;
            }
        }
        //Call AD and return user groups for user
        public static string GetGroups(string username)
        {
            string[] output = null;
            string result = null;
            try
            {
                using (var ctx = new PrincipalContext(ContextType.Domain))
                using (var user = UserPrincipal.FindByIdentity(ctx, username))
                {
                    if (user != null)
                    {
                        output = user.GetGroups() //this returns a collection of principal objects
                            .Select(x => x.SamAccountName) // select the name.  you may change this to choose the display name or whatever you want
                            .ToArray(); // convert to string array
                    }
                }
            }
            catch (Exception ex)
            {
                result = ex.ToString();
            }
            return result;
        }

        //
        //!--App Info--!!
        //

        //Uses registry entries to populate listbox
        public void GetInstalledApps()
        {
            string uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(uninstallKey))
            {
                foreach (string skName in rk.GetSubKeyNames())
                {
                    using (RegistryKey sk = rk.OpenSubKey(skName))
                    {
                        try
                        {
                            AppsList.Items.Add(sk.GetValue("DisplayName"));
                        }
                        catch (Exception ex)
                        {
                            AppsList.Items.Add("Unable to obtain");
                        }
                    }
                }
                //Label1.Content = AppsList.Items.Count.ToString();
            }
        }

        public MainWindow()
        {

            InitializeComponent();
            //--System Tab Start--
            //Build Information
            CompNameVal.Content = ComputerName();
            AssetNumVal.Content = GetAssetTag();
            BuildDateTimeVal.Content = InstallDate();
            InitBuildVerVal.Content = "TODO"; //Probably use a filedrop or call the old registry key value
            CurBuildVerVal.Content = BuildNumber();
            //System Software
            OperatingSysVal.Content = GetOSFriendlyName();
            KernelVerVal.Content = Kernel();
            IEVerVal.Content = IEVersion(); //Maybe registry?
            MSOfficeVerVal.Content = MSOfficeVersion();
            AppVVerVal.Content = AppVVersion();
            //--System Tab End--

            //--Hardware Tab Start--
            //Hardware Information
            MakeVal.Content = MakeName();
            ModelVal.Content = ModelName();
            UEFIVerVal.Content = UEFIBIOSVersion();
            SKUNumVal.Content = SKUNumber();
            ProcessorVal.Content = ProcessorName();
            MemoryVal.Content = PhysicalMemory();
            //Networking Information
            IPAddrVal.Content = IPAddress();
            MACAddrVal.Content = MACAddress();
            GatewayVal.Content = GatewayAddress();
            DNSServersVal.Content = DnsAddress();
            //--Hardware Tab End--

            //--Roles Tab Start--
            //Domain Information
            CurDomainVal.Content = DomainName();
            UsernameVal.Content = UserName();
            ADSiteVal.Content = ADSiteName();
            DCServerVal.Content = DCServerName();
            SubAreaIDVal.Content = SubAreaName();
            ADGroupsVal.Items.Add(GetGroups(System.Security.Principal.WindowsIdentity.GetCurrent().Name));
            //--Roles Tab End--
        }

        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Close_Button_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown(); //programatically close the app
        }

        private void Export_Button_Click(object sender, RoutedEventArgs e)
        {

            using (XmlWriter writer = XmlWriter.Create("C:\\Users\\Jack\\Desktop\\report.xml"))
            {
                writer.WriteStartElement(ComputerName());
                writer.WriteStartElement("System");
                writer.WriteElementString("ITNumber", ComputerName());
                writer.WriteElementString("AssetTag", GetAssetTag());
                writer.WriteElementString("InstallDate", InstallDate());
                writer.WriteElementString("InitialBuild", "Not yet implemented");
                writer.WriteElementString("CurrentBuild", BuildNumber());
                writer.WriteStartElement("SystemSoftware");
                writer.WriteElementString("OperatingSystem", GetOSFriendlyName());
                writer.WriteElementString("KernelVersion", Kernel());
                writer.WriteElementString("IEVersion", IEVersion());
                writer.WriteElementString("OfficeVersion", MSOfficeVersion());
                writer.WriteElementString("AppVVersion", AppVVersion());
                writer.WriteEndElement();
                writer.WriteStartElement("Hardware");
                writer.WriteElementString("Make",MakeName());
                writer.WriteElementString("Model",ModelName());
                writer.WriteElementString("UEFIVersion",UEFIBIOSVersion());
                writer.WriteElementString("SKU",SKUNumber());
                writer.WriteElementString("Processor",ProcessorName());
                writer.WriteElementString("RAM",PhysicalMemory());
                writer.WriteStartElement("Network");
                writer.WriteElementString("IPAddress",IPAddress());
                writer.WriteElementString("MACAddress",MACAddress());
                writer.WriteElementString("GatewayAddress", GatewayAddress());
                writer.WriteElementString("DNSAddress", DnsAddress());
                writer.WriteStartElement("Domain");
                writer.WriteElementString("DomainName", DomainName());
                writer.WriteElementString("UserName", UserName());
                writer.WriteElementString("SiteName", ADSiteName());
                writer.WriteElementString("DCServerName", DCServerName());
                writer.WriteElementString("SubAreaName", SubAreaName());
                writer.WriteElementString("ADGroups", ADGroupsVal.Items.Add(GetGroups(System.Security.Principal.WindowsIdentity.GetCurrent().Name)).ToString());
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.Flush();
            }
            MessageBoxResult result = MessageBox.Show("File has been exported to the Desktop.",
                              "Confirmation",
                              MessageBoxButton.OK,
                              MessageBoxImage.Information);
            if (result == MessageBoxResult.Yes)
            {
                Application.Current.Shutdown();
            }
        }

        private void AppsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            InitializeComponent();
            //Call this GetInstalledApps in Window Forms Constructor.  
            GetInstalledApps();
        }
    }
}

