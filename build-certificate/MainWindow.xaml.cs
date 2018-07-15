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

namespace build_certificate
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //Calls WMI and searches for OS caption name
        public static string GetOSFriendlyName()
        {
            string result = string.Empty;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Caption FROM Win32_OperatingSystem");
            foreach (ManagementObject os in searcher.Get())
            {
                result = os["Caption"].ToString();
                break;
            }
            return result;
        }
        //Calls WMI and searches for SMBIOSAssetTag
        public static string GetAssetTag()
        {
            string result = string.Empty;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_SystemEnclosure");
            foreach (ManagementObject SMBIOSAssetTag in searcher.Get())
            {
                result = SMBIOSAssetTag["SMBIOSAssetTag"].ToString();
                break;
            }
            return result;
        }
        //Calls WMI and searches for InstallDate
        public static string InstallDate()
        {
            string result = string.Empty;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
            foreach (ManagementObject InstallDate in searcher.Get())
            {
                result = InstallDate["InstallDate"].ToString();
                break;
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
                result = build.Insert(0, "Build "); //Insert "Build" at the first index
                try //now we check the registry key values for the Release ID
                {
                    using (RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Windows NT\\CurrentVersion"))
                    {
                        if (key != null)
                        {
                            Object o = key.GetValue("ReleaseID");
                            if (o != null)
                            {
                                version =  o.ToString(); //convert object to string
                            }
                        }
                    }
                }
                catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                {
                    //react appropriately
                }
                result = version.Insert(0, "Version "); //Insert "Version" at the first index

                break;
            }
            return result;
        }
        //Calls WMI and searches for Kernel Version
        public static string Kernel()
        {
            string result = string.Empty;

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
            foreach (ManagementObject Kernel in searcher.Get())
            {
                result = Kernel["Version"].ToString();
                break;
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

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
            foreach (ManagementObject MakeName in searcher.Get())
            {
                result = MakeName["Manufacturer"].ToString();
                break;
            }
            return result;
        }
        //Calls WMI and searches for Model
        public static string ModelName()
        {
            string result = string.Empty;

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
            foreach (ManagementObject ModelName in searcher.Get())
            {
                result = ModelName["Model"].ToString();
                break;
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
                result = UEFIBIOSVersion["BIOSVersion"].ToString();
                break;
            }
            return result;
        }
        //Calls WMI and searches for SystemSKU
        public static string SKUNumber()
        {
            string result = string.Empty;

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
            foreach (ManagementObject SKUNumber in searcher.Get())
            {
                result = SKUNumber["SystemSkuNumber"].ToString();
                break;
            }
            return result;
        }
        //Calls WMI and searches for Name
        public static string ProcessorName()
        {
            string result = string.Empty;

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
            foreach (ManagementObject ProcessorName in searcher.Get())
            {
                result = ProcessorName["Name"].ToString();
                break;
            }
            return result;
        }
        //Calls WMI and searches for TotalVisibleMemorySize
        public static string PhysicalMemory()
        {
            double memval;
            double result = 0;

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
            foreach (ManagementObject PhysicalMemory in searcher.Get())
            {
                memval = Convert.ToDouble(PhysicalMemory["TotalVisibleMemorySize"]);
                result = Math.Round((memval / (1024 * 1024)), 0);
                break;
            }
            return result.ToString(result+" GB");
        }
        //Calls DNS class to return IP for hostname in Index 1 (Index 0 is normally loopback)
        public static string IPAddress()
        {
            string hostName = Dns.GetHostName();
            string result = Dns.GetHostEntry(hostName).AddressList[1].ToString();

            return result;
        }
        //Calls Physical Address Class to return MAC Address
        public static string MACAddress()
        {
            string result = string.Empty;                                                           //
            IPGlobalProperties computerProperties = IPGlobalProperties.GetIPGlobalProperties();     // Crappy method to get NIC MAC Address
            NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();                   // Probably better ways...
            foreach (NetworkInterface adapter in nics)                                              //
            {
                IPInterfaceProperties properties = adapter.GetIPProperties();
                result = adapter.GetPhysicalAddress().ToString();
                break;

            }
            result = System.Text.RegularExpressions.Regex.Replace(result.ToString(), ".{2}", "$0-"); //replace with hyphens every 2nd character
            result = result.Remove(result.Length - 1); //trim the extra '-'
            return result;
        }

        //Calls IPAddress class and looks for first network card and grabs it's gateway address
        //public static IPAddress GatewayAddress()
        //{
        //    IPAddress result = null;
        //    var cards = NetworkInterface.GetAllNetworkInterfaces().ToList();
        //    if (cards.Any())
        //    {
        //        foreach (var card in cards)
        //        {
        //            var props = card.GetIPProperties();
        //            if (props == null)
        //                continue;

        //            var gateways = props.GatewayAddresses;
        //            if (!gateways.Any())
        //                continue;

        //            var gateway =
        //                gateways.FirstOrDefault(g => g.Address.AddressFamily.ToString() == "InterNetwork");
        //            if (gateway == null)
        //                continue;

        //            result = gateway.Address;
        //            break;
        //        };
        //    }
        //                return result;
        //}
        //Calls WMI and searches for DomainName
        public static string GatewayAddress()
        {
            string result = string.Empty;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE INDEX=1");
            foreach (ManagementObject GatewayAddress in searcher.Get())
            {
                try
                {
                    result = GatewayAddress["DefaultIPGateway"].ToString();
                }
                catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                {
                    //react appropriately
                    if (result == null)
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
            NetworkInterface[] adapters = NetworkInterface.GetAllNetworkInterfaces();
            foreach (NetworkInterface adapter in adapters)
            {

                IPInterfaceProperties adapterProperties = adapter.GetIPProperties();
                IPAddressCollection dnsServers = adapterProperties.DnsAddresses;
                if (dnsServers.Count > 0)
                {
                    Console.WriteLine(adapter.Description);
                    foreach (IPAddress dns in dnsServers)
                    {
                            
                        result = dns.ToString();
                        break;

                    }
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
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_NTDomain");
            foreach (ManagementObject DomainName in searcher.Get())
            {
                try
                {
                    result = DomainName["DomainName"].ToString();
                }
                catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                {
                    //react appropriately
                    if (result == null)
                        break;
                }

                break;
            }
            return result;
        }
        //Calls WMI and searches for DomainName
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
                        break;
                }

                break;
            }
            return result;
        }
        //Calls WMI and searches for DomainName
        public static string ADSiteName()
        {
            string result = string.Empty;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_NTDomain");
            foreach (ManagementObject ADSiteName in searcher.Get())
            {
                try
                {
                    result = ADSiteName["ClientSiteName"].ToString();
                }
                catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                {
                    //react appropriately
                    if (result == null)
                        break;
                }

                break;
            }
            return result;
        }

        //Call AD and return user groups for user
        public static string[] GetGroups(string username)
        {
            string[] output = null;

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

            return output;
        }

        public MainWindow()
        {

            InitializeComponent();
            //--System Tab Start--
            //Build Information
            CompNameVal.Content = Environment.MachineName; //We can get this item locally in .NET
            AssetNumVal.Content = GetAssetTag();
            BuildDateTimeVal.Content = InstallDate();
            InitBuildVerVal.Content = "TODO"; //Probably use a filedrop or call the old registry key value
            CurBuildVerVal.Content = BuildNumber();
            //System Software
            OperatingSysVal.Content = GetOSFriendlyName();
            KernelVerVal.Content = Kernel();
            IEVerVal.Content = "TODO"; //Maybe registry?
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
            //SubAreaIDVal.Content = GetGroups("Jack");
            //--Roles Tab End--
        }

        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }


    }
}

