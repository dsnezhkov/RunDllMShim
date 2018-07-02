using System;
using System.Management;

namespace KlassLibrary
{
    public class Klass
    {

        private static void WmiNetQ()
        {

            string query = "SELECT * FROM Win32_NetworkAdapterConfiguration"
                 + " WHERE IPEnabled = 'TRUE'";

            ManagementObjectSearcher moSearch = new ManagementObjectSearcher(query);
            ManagementObjectCollection moCollection = moSearch.Get();

            foreach (ManagementObject mo in moCollection)
            {
                Console.WriteLine("HostName = " + mo["DNSHostName"]);
                Console.WriteLine("Description = " + mo["Description"]);


                string[] addresses = (string[])mo["IPAddress"];
                foreach (string ipaddress in addresses)
                {
                    Console.WriteLine("IPAddress = " + ipaddress);
                }

                string[] subnets = (string[])mo["IPSubnet"];
                foreach (string ipsubnet in subnets)
                {
                    Console.WriteLine("IPSubnet = " + ipsubnet);
                }


                string[] defaultgateways = (string[])mo["DefaultIPGateway"];
                foreach (string defaultipgateway in defaultgateways)
                {
                    Console.WriteLine("DefaultIPGateway = " + defaultipgateway);
                }
            }

        }
        public static int RunCode(string arg)
        {
            Console.WriteLine("== Default static class method in managed code");
            WmiNetQ();
            return 0;
        }

        public static int SKlassMethod(string arg)
        {
            Console.WriteLine("Static class method in managed code.");
            switch (arg)
            {
                case "GetNetwork":
                    Console.WriteLine("Invoking WMI feature for network query");
                    WmiNetQ();
                    break;
                default:
                    Console.WriteLine("Not sure what you are asking for.");
                    break;
            }

            return 0;
        }

    }
}
