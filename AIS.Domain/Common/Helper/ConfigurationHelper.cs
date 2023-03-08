using System.Configuration;

namespace AIS.Domain.Common.Helper
{
    public class ConfigurationHelper
    {
        public static string GetValueConfig(string key)
        {
            return ConfigurationManager.AppSettings[key];
        }
    }
}
