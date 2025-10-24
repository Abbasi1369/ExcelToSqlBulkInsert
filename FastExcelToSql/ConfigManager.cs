using System.Text.Json;

//using static ExcelToSqlBulkLibrary.ExcelToSqlBulkLibrary.ExcelToSqlBulkProcessor;
using static ExcelToSqlBulkLibrary.ExcelToSqlBulkProcessor;

namespace ExcelToSql
{
    public static class ConfigManager
    {
        private static readonly string ConfigFileName = "excel-to-sql-config.json";

        public static ConfigModel LoadConfig()
        {
            try
            {
                if (!File.Exists(ConfigFileName))
                {
                    CreateDefaultConfig();
                    MessageBox.Show(
                        $"فایل تنظیمات ایجاد شد: {ConfigFileName}\nلطفاً فایل تنظیمات را پر کنید و دوباره برنامه را اجرا کنید.",
                        "فایل تنظیمات",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    Application.Exit();
                    return null;
                }

                var json = File.ReadAllText(ConfigFileName);
                var config = JsonSerializer.Deserialize<ConfigModel>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

                ValidateConfig(config);
                return config;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"خطا در خواندن فایل تنظیمات: {ex.Message}",
                    "خطا",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return null;
            }
        }

        private static void CreateDefaultConfig()
        {
            var defaultConfig = new ConfigModel
            {
                InputPath = "C:\\ExcelFiles",
                ConnectionString = "Server=.;Database=TestDB;Integrated Security=true;TrustServerCertificate=true",
                TableName = "",
                UseAppendLogic = true,
                DetectColumnTypes = true,
                OverwriteTable = true,
                ParallelProcessing = true,
                MaxDegreeOfParallelism = 3,
                NamingStrategy = "Clean",
                EnableLogging = true,
                Verbose = false
            };

            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(defaultConfig, options);
            File.WriteAllText(ConfigFileName, json);
        }

        private static void ValidateConfig(ConfigModel config)
        {
            if (string.IsNullOrEmpty(config.ConnectionString))
                throw new ArgumentException("ConnectionString در فایل تنظیمات خالی است");
        }

        public static ColumnNamingStrategy ParseNamingStrategy(string strategy)
        {
            return strategy?.ToLower() switch
            {
                "original" => ColumnNamingStrategy.Original,
                "clean" => ColumnNamingStrategy.Clean,
                "brackets" => ColumnNamingStrategy.Brackets,
                _ => ColumnNamingStrategy.Clean
            };
        }
    }
}
