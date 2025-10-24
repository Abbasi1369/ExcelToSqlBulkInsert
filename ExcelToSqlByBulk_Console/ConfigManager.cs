using System.Text.Json;
using System.Text.Json.Serialization;
using static ExcelToSqlBulkLibrary.ExcelToSqlBulkProcessor;

namespace ExcelToSqlByBulk_Console
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
                    Console.WriteLine($"📁 ConfigFile Start: {ConfigFileName}");
                    //Console.WriteLine("⚠️ لطفاً فایل تنظیمات را پر کنید و دوباره اجرا کنید.");
                    Environment.Exit(0);
                }

                var json = File.ReadAllText(ConfigFileName);
                var config = JsonSerializer.Deserialize<ConfigModel>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true,
                    ReadCommentHandling = JsonCommentHandling.Skip
                });

                ValidateConfig(config);
                return config;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error In ConfigFile: {ex.Message}");
                Environment.Exit(1);
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

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            };

            var json = JsonSerializer.Serialize(defaultConfig, options);
            File.WriteAllText(ConfigFileName, json);
        }

        private static void ValidateConfig(ConfigModel config)
        {
            if (string.IsNullOrEmpty(config.InputPath))
                throw new ArgumentException("InputPath Empty in config");

            if (string.IsNullOrEmpty(config.ConnectionString))
                throw new ArgumentException("ConnectionString Empty in config");

            if (!Directory.Exists(config.InputPath) && !File.Exists(config.InputPath))
                throw new DirectoryNotFoundException($"DirectoryNotFoundException: {config.InputPath}");

            if (config.MaxDegreeOfParallelism < 1 || config.MaxDegreeOfParallelism > 10)
                throw new ArgumentException("MaxDegreeOfParallelism  Min :1 And Max 10");
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
