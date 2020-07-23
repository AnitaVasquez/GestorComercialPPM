using Newtonsoft.Json;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Helpers
{
    public class Tools
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public static string ReadSetting(string key)
        {
            try
            {
                var appSettings = ConfigurationManager.AppSettings;
                string result = appSettings[key] ?? "Not Found";
                return result.ToString();
            }
            catch (ConfigurationErrorsException ex)
            {
                logger.Error(ex, "Stopped program because of exception");
                return string.Empty;
            }
        } 

        public static string GetJson(string texto)
        {
            var res = "";
            try
            {
                var json = JsonConvert.DeserializeObject<dynamic>(texto);
                res = json["access_token"];
                return res;
            }
            catch (System.Exception ex)
            {
                logger.Error(ex, "Stopped program because of exception");
                return res;
            }
        }
    }
}