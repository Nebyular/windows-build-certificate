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
        //Calls registry and searches for Iniital install version
        public static string InitialInstallVersion()
        {
            string result = string.Empty;
            try //now we check the registry key values for the Release ID
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey("System\\Select\\DHS"))
                {
                    if (key != null)
                        if (key != null)
                        {
                            Object o = key.GetValue("BuildVersion");
                            if (o != null)
                            {
                                result = o.ToString();  //convert object to string
                            }
                        }
                    result = "Unable to obtain";
                }
            }
            catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
            {
                result = "Unable to obtain";
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
                                result = "Version " + version;
                            }
                        }
                    }
                }
                catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                {
                    result = "Unable to obtain";
                }

                break;
            }
            return result;
        }
        //Calls WMI and searches for Kernel Version
        public static string Kernel()
        {
            string result = string.Empty;
            string version = string.Empty;
            string ubr = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Version FROM Win32_OperatingSystem");
                foreach (ManagementObject Kernel in searcher.Get())
                {
                    version = Kernel["Version"].ToString();
                    break;
                }
            }
            catch (Exception ex)
            {
                result = "Unable to obtain";
            }
            try
            {
                using (RegistryKey key2 = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Windows NT\\CurrentVersion"))
                {
                    if (key2 != null)
                    {
                        Object o = key2.GetValue("UBR");
                        if (o != null)
                        {
                            ubr = o.ToString(); //convert object to string
                        }
                    }
                }
            }
            catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
            {
                result = "Unable to obtain";
            }
            result = version + "." + ubr;
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
        //Calls WMI and searches for BIOS UUID
        public static string UUID()
        {
            string result = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT UUID FROM WIN32_COMPUTERSYSTEMPRODUCT");
                foreach (ManagementObject UUID in searcher.Get())
                {
                    result = UUID["UUID"].ToString();
                    break;
                }
            }
            catch (Exception ex)
            {
                result = "Unable to obtain";
            }
            return result;
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
            string result = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE INDEX=2");
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


        public MainWindow()
        {

            InitializeComponent();
            //--System Tab Start--
            //Build Information
            CompNameVal.Content = ComputerName();
            AssetNumVal.Content = GetAssetTag();
            BuildDateTimeVal.Content = InstallDate();
            InitBuildVerVal.Content = InitialInstallVersion(); //Probably use a filedrop or call the old registry key value
            CurBuildVerVal.Content = BuildNumber();
            //System Software
            OperatingSysVal.Content = GetOSFriendlyName();
            KernelVerVal.Content = Kernel();
            //--System Tab End--

            //--Hardware Tab Start--
            //Hardware Information
            MakeVal.Content = MakeName();
            ModelVal.Content = ModelName();
            UEFIVerVal.Content = UEFIBIOSVersion();
            SKUNumVal.Content = SKUNumber();
            ProcessorVal.Content = ProcessorName();
            MemoryVal.Content = PhysicalMemory();
            BIOSUUIDVal.Content = UUID();
            //Networking Information
            IPAddrVal.Content = IPAddress();
            MACAddrVal.Content = MACAddress();
            //--Hardware Tab End--

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
                writer.WriteEndElement();
                writer.WriteStartElement("Hardware");
                writer.WriteElementString("Make",MakeName());
                writer.WriteElementString("Model",ModelName());
                writer.WriteElementString("UEFIVersion",UEFIBIOSVersion());
                writer.WriteElementString("SKU",SKUNumber());
                writer.WriteElementString("Processor",ProcessorName());
                writer.WriteElementString("RAM",PhysicalMemory());
                writer.WriteElementString("BIOSUUID", UUID());
                writer.WriteEndElement();
                writer.WriteStartElement("Network");
                writer.WriteElementString("IPAddress",IPAddress());
                writer.WriteElementString("MACAddress",MACAddress());
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
    }
}

