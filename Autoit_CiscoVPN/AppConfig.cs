using System.Configuration;

namespace Autoit_CiscoVPN
{
    public class AppConfig
    {
        public static string userName = ConfigurationManager.AppSettings["username"];
        public static string passWord = ConfigurationManager.AppSettings["password"];
        public static string ciscoExePath = ConfigurationManager.AppSettings["ciscoexepath"];
        public static string domainName = ConfigurationManager.AppSettings["domain"];
        public static string groupName = ConfigurationManager.AppSettings["group"];
        public static string checkcertificationerror = ConfigurationManager.AppSettings["checkcertificationerror"];

    }
}
