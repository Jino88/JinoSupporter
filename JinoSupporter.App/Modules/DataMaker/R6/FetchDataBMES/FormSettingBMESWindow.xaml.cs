using Newtonsoft.Json;
using System;
using System.IO;
using System.Windows;
using WorkbenchHost.Infrastructure;

namespace DataMaker.R6.FetchDataBMES
{
    public partial class FormSettingBMESWindow : Window
    {
        public static string AppDataDirectory => AppSettingsPathManager.GetModuleDirectory("DataMaker");
        public static string DefaultDataFilePath => AppSettingsPathManager.GetModuleFilePath("DataMaker", "BMESIDINFO.json");
        public static string LastPathMarkerFilePath => WorkbenchSettingsStore.SettingsFilePath;

        private static string LastUsedFilePath = ResolveInitialDataFilePath();
        public static InfoID? infoID;

        public FormSettingBMESWindow()
        {
            InitializeComponent();

            infoID = EnsureLoadedInfo();
            CT_TB_ID.Text = infoID.LoginID;
            CT_TB_PASSWORD.Password = infoID.Password;
        }

        private void CT_BT_SAVE_Click(object sender, RoutedEventArgs e)
        {
            infoID = new InfoID(CT_TB_ID.Text, CT_TB_PASSWORD.Password);
            if (!SaveToUnifiedSettings(infoID))
            {
                MessageBox.Show("Failed to save BMES credentials.", "Save Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            MessageBox.Show(
                $"Credentials were saved.\n\nFile: {Path.GetFileName(WorkbenchSettingsStore.SettingsFilePath)}",
                "Saved",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void CT_BT_LOAD_Click(object sender, RoutedEventArgs e)
        {
            InfoID? loadedInfo = LoadInfoFromCentralSettings();

            if (loadedInfo is null)
            {
                MessageBox.Show(
                    $"Failed to load credentials.\n\nFile: {Path.GetFileName(WorkbenchSettingsStore.SettingsFilePath)}",
                    "Load Failed",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            infoID = loadedInfo;
            CT_TB_ID.Text = infoID.LoginID;
            CT_TB_PASSWORD.Password = infoID.Password;

            MessageBox.Show(
                $"Credentials were loaded.\n\nFile: {Path.GetFileName(WorkbenchSettingsStore.SettingsFilePath)}",
                "Loaded",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private static bool SaveDataToPath(string id, string password, string filePath)
        {
            try
            {
                string normalizedPath = NormalizePath(filePath);
                string directory = Path.GetDirectoryName(normalizedPath) ?? AppDataDirectory;
                Directory.CreateDirectory(directory);

                var payload = new InfoID(id, password);
                string jsonString = JsonConvert.SerializeObject(payload, Formatting.Indented);
                File.WriteAllText(normalizedPath, jsonString);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static InfoID? LoadDataFromPath(string filePath)
        {
            string normalizedPath = NormalizePath(filePath);
            if (!File.Exists(normalizedPath))
            {
                return null;
            }

            try
            {
                string jsonString = File.ReadAllText(normalizedPath);
                return JsonConvert.DeserializeObject<InfoID>(jsonString);
            }
            catch
            {
                return null;
            }
        }

        public static bool SaveData(string id, string password)
        {
            infoID = new InfoID(id, password);
            return SaveToUnifiedSettings(infoID);
        }

        public static InfoID? LoadData()
        {
            infoID = LoadInfoFromCentralSettings() ?? LoadDataFromPath(LastUsedFilePath);
            return infoID;
        }

        public static string GetCurrentDataFilePath()
        {
            return WorkbenchSettingsStore.SettingsFilePath;
        }

        public static void SetCurrentDataFilePath(string filePath)
        {
            LastUsedFilePath = NormalizePath(filePath);
            WorkbenchSettingsStore.UpdateSettings(settings => settings.Bmes.LastImportExportFilePath = LastUsedFilePath);
        }

        public static InfoID EnsureLoadedInfo()
        {
            if (infoID is null)
            {
                infoID = LoadInfoFromCentralSettings() ?? LoadDataFromPath(LastUsedFilePath);
            }

            if (infoID is null &&
                !string.Equals(LastUsedFilePath, DefaultDataFilePath, StringComparison.OrdinalIgnoreCase))
            {
                infoID = LoadDataFromPath(DefaultDataFilePath);
                if (infoID is not null)
                {
                    LastUsedFilePath = DefaultDataFilePath;
                    SaveInfoToCentralSettings(infoID, LastUsedFilePath);
                }
            }

            if (infoID is null)
            {
                infoID = new InfoID("", "");
            }

            return infoID;
        }

        private static string ResolveInitialDataFilePath()
        {
            try
            {
                string savedPath = WorkbenchSettingsStore.GetSettings().Bmes.LastImportExportFilePath;
                if (!string.IsNullOrWhiteSpace(savedPath))
                {
                    return NormalizePath(savedPath);
                }
            }
            catch
            {
                // Ignore and use default path.
            }

            return DefaultDataFilePath;
        }

        private static void PersistLastUsedFilePath(string filePath)
        {
            WorkbenchSettingsStore.UpdateSettings(settings => settings.Bmes.LastImportExportFilePath = NormalizePath(filePath));
        }

        private static string NormalizePath(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return DefaultDataFilePath;
            }

            return Path.GetFullPath(filePath);
        }

        private static InfoID? LoadInfoFromCentralSettings()
        {
            try
            {
                WorkbenchSettingsStore.BmesSetting bmes = WorkbenchSettingsStore.GetSettings().Bmes;
                if (string.IsNullOrWhiteSpace(bmes.LoginId) && string.IsNullOrWhiteSpace(bmes.Password))
                {
                    return null;
                }

                if (!string.IsNullOrWhiteSpace(bmes.LastImportExportFilePath))
                {
                    LastUsedFilePath = NormalizePath(bmes.LastImportExportFilePath);
                }

                return new InfoID(bmes.LoginId, bmes.Password);
            }
            catch
            {
                return null;
            }
        }

        private static void SaveInfoToCentralSettings(InfoID currentInfo, string filePath)
        {
            WorkbenchSettingsStore.UpdateSettings(settings =>
            {
                settings.Bmes.LoginId = currentInfo.LoginID;
                settings.Bmes.Password = currentInfo.Password;
                settings.Bmes.LastImportExportFilePath = NormalizePath(filePath);
            });
        }

        private static bool SaveToUnifiedSettings(InfoID currentInfo)
        {
            try
            {
                SaveInfoToCentralSettings(currentInfo, LastUsedFilePath);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

    public class InfoID
    {
        public string LoginID { get; set; }
        public string Password { get; set; }

        public InfoID(string loginID, string password)
        {
            LoginID = loginID;
            Password = password;
        }
    }
}
