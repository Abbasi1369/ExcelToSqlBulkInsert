using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace ExcelToSql
{
   
    
public class ConfigModel
    {
        [JsonPropertyName("inputPath")]
        public string InputPath { get; set; } = "";

        [JsonPropertyName("connectionString")]
        public string ConnectionString { get; set; } = "";

        [JsonPropertyName("tableName")]
        public string TableName { get; set; } = "";

        [JsonPropertyName("useAppendLogic")]
        public bool UseAppendLogic { get; set; } = true;

        [JsonPropertyName("detectColumnTypes")]
        public bool DetectColumnTypes { get; set; } = true;

        [JsonPropertyName("overwriteTable")]
        public bool OverwriteTable { get; set; } = true;

        [JsonPropertyName("parallelProcessing")]
        public bool ParallelProcessing { get; set; } = true;

        [JsonPropertyName("maxDegreeOfParallelism")]
        public int MaxDegreeOfParallelism { get; set; } = 3;

        [JsonPropertyName("namingStrategy")]
        public string NamingStrategy { get; set; } = "Clean";

        [JsonPropertyName("enableLogging")]
        public bool EnableLogging { get; set; } = true;

        [JsonPropertyName("verbose")]
        public bool Verbose { get; set; } = false;
    }
}
