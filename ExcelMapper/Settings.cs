using System;
using System.Configuration;
using NLog;
using NLog.Internal;

namespace ExcelMapper
{
    public class Settings
    {
        public static readonly string MappingJSONPathSetting = "MappingJSONPath";
        public static readonly string LastRunSetting = "LastRun";

        private readonly Logger logger = LogManager.GetCurrentClassLogger();
        private Configuration config = null;

        public Settings()
        {
            try
            {
                config = System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                logger.Info("Reading app settings from: " + config.FilePath);

                LastRun = DateTime.Now;
                string value = ReadSetting(LastRunSetting);
                if (!string.IsNullOrEmpty(value))
                {
                    LastRun = Convert.ToDateTime(value);
                }

                MappingJSONPath = ReadSetting(MappingJSONPathSetting);
            }
            catch (FormatException e)
            {
                logger.Error("Unable to convert app setting (" + e + ")");
                Environment.Exit(-1);
            }
            catch (Exception e)
            {
                logger.Error("Unable to correctly read app settings (" + e + ")");
                Environment.Exit(-1);
            }
        }

        public DateTime LastRun { get; set; }
        public string MappingJSONPath { get; set; }

        public void WriteSetting(string key, string value)
        {
            try
            {
                var settings = config.AppSettings.Settings;
                if (settings[key] == null)
                {
                    settings.Add(key, value);
                }
                else
                {
                    settings[key].Value = value;
                }
                config.Save(ConfigurationSaveMode.Modified);
                System.Configuration.ConfigurationManager.RefreshSection(config.AppSettings.SectionInformation.Name);
            }
            catch (ConfigurationErrorsException e)
            {
                logger.Error("Unable to write to the app settings (" + e + ").");
            }
        }

        private string ReadSetting(string key)
        {
            string result = System.Configuration.ConfigurationManager.AppSettings[key];
            if (result == null)
            {
                throw new ArgumentException(key);
            }
            return result;
        }
    }
}
