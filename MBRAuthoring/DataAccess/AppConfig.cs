using MBRAuthoring.Utilities;
using System.Configuration;

namespace MBRAuthoring.DataAccess
{
    public class AppConfig
    {
        public static string ConnString
        {
            get
            {
                //return ConfigurationManager.ConnectionStrings[ConnectionStringName].ConnectionString;
                return GenericUtility.DecryptText(ConfigurationManager.ConnectionStrings[ConnectionStringName].ConnectionString);
            }
        }

        public static string ConnectionStringName
        {
            get
            {
                return ConfigurationManager.AppSettings["Environment"] + "_" + ConfigurationManager.AppSettings["ConnectionStringName"];
            }
        }
    }
}
