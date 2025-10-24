using ExcelDataReader;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;




namespace ExcelToSqlBulkLibrary
{
    public class ExcelToSqlBulkProcessor
    {


        public enum ColumnNamingStrategy
        {
            Original,    // نام اصلی ستون‌ها
            Clean,       // پاکسازی کاراکترهای غیرمجاز
            Brackets     // استفاده از براکت برای نام‌های خاص
        }

        private readonly string _connectionString;
        private readonly ColumnNamingStrategy _namingStrategy;

        public ExcelToSqlBulkProcessor(string connectionString, ColumnNamingStrategy namingStrategy = ColumnNamingStrategy.Clean)
        {
            _connectionString = EnsureTrustServerCertificate(connectionString);
            _namingStrategy = namingStrategy;
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        private string EnsureTrustServerCertificate(string connectionString)
        {
            if (!connectionString.Contains("TrustServerCertificate", StringComparison.OrdinalIgnoreCase))
            {
                if (connectionString.EndsWith(";"))
                    return connectionString + "TrustServerCertificate=true;";
                else
                    return connectionString + ";TrustServerCertificate=true;";
            }
            return connectionString;
        }

        public async Task ProcessExcelFileWithTableCreation(string filePath, string tableName,
        bool detectColumnTypes = true, bool overwriteTable = true, bool forceCreate = false)
        {
            // خواندن فایل Excel
            var dataTable = ReadExcelToDataTable(filePath);

            // تشخیص و تبدیل انواع ستون‌ها
            if (detectColumnTypes)
            {
                dataTable = DetectAndConvertColumnTypes(dataTable);
            }

            // پاکسازی نام جدول
            tableName = CleanTableName(tableName);

            // اگر forceCreate=true یا جدول وجود ندارد، ایجاد کن
            bool tableAlreadyExists = await TableExists(tableName);


            if (forceCreate || !tableAlreadyExists)
            {
                // ایجاد یا بازنویسی جدول
                await CreateOrOverwriteTable(dataTable, tableName, overwriteTable);
            }
            else if (tableAlreadyExists && overwriteTable)
            {
                // فقط اگر overwriteTable=true باشه، جدول رو بازنویسی می‌کنه
                await CreateOrOverwriteTable(dataTable, tableName, true);
            }

            // درج داده‌ها
            await BulkInsertToSqlServer(dataTable, tableName);
        }
       
        private DataTable ReadExcelToDataTable(string filePath)
        {

           

            //// برای فایل‌های xlsx از OpenXml استفاده کن (سریع‌تر)
            //if ( filePath.EndsWith(".xlsx"))
            //{
            //    try
            //    {
            //        return ReadExcelWithOpenXml(filePath);
            //    }
            //    catch
            //    {
            //        // اگر OpenXml خطا داد، به روش معمول برگرد
            //        Console.WriteLine("⚠️ OpenXml failed, using ExcelDataReader");
            //    }
            //}
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            return dataSet.Tables[0];
        }

        private DataTable DetectAndConvertColumnTypes(DataTable rawDataTable)
        {
            var resultTable = new DataTable();

            // ایجاد ستون‌ها با انواع تشخیص داده شده
            foreach (DataColumn rawColumn in rawDataTable.Columns)
            {
                var columnName = CleanColumnNameAccordingToStrategy(rawColumn.ColumnName);
                var detectedType = DetectColumnType(rawDataTable, rawColumn.ColumnName);
                resultTable.Columns.Add(columnName, detectedType);
            }

            // کپی داده‌ها با تبدیل نوع
            foreach (DataRow rawRow in rawDataTable.Rows)
            {
                var newRow = resultTable.NewRow();

                for (int i = 0; i < resultTable.Columns.Count; i++)
                {
                    var column = resultTable.Columns[i];
                    var rawValue = rawRow[i];

                    newRow[i] = SafeConvertValue(rawValue, column.DataType);
                }

                resultTable.Rows.Add(newRow);
            }

            return resultTable;
        }

        private string CleanColumnNameAccordingToStrategy(string columnName)
        {
            return _namingStrategy switch
            {
                ColumnNamingStrategy.Original => columnName,
                ColumnNamingStrategy.Clean => CleanColumnName(columnName, false),
                ColumnNamingStrategy.Brackets => CleanColumnName(columnName, true),
                _ => CleanColumnName(columnName, false)
            };
        }
        private Type DetectColumnType(DataTable dataTable, string columnName)
        {
            var samples = dataTable.AsEnumerable()
                .Take(1000)
                .Select(row => row[columnName]?.ToString())
                .Where(val => !string.IsNullOrEmpty(val))
                .ToList();

            if (!samples.Any())
                return typeof(string);
            if (samples.Any(c => c.StartsWith("0") && c!="0")) return typeof(string);

            int intCount = 0, doubleCount = 0, dateCount = 0, boolCount = 0, stringCount = 0;

            foreach (var sample in samples)
            {
                if (int.TryParse(sample, out _)) intCount++;
                if (double.TryParse(sample, out _)) doubleCount++;
                if (DateTime.TryParse(sample, out _)) dateCount++;
                if (bool.TryParse(sample, out _) || sample == "1" || sample == "0") boolCount++;
                else stringCount++;
            }

            var total = samples.Count;
            // اگر بیش از ۲٪ متن غیرقابل تبدیل داریم، string باشه
            if (stringCount > total * 0.02) return typeof(string);

            if (dateCount > total * 0.9) return typeof(DateTime);
            if (boolCount > total * 0.9) return typeof(bool);
            if (doubleCount > total * 0.9 && doubleCount > intCount) return typeof(double);
            if (intCount > total * 0.9) return typeof(int);

            return typeof(string);
        }

        private object SafeConvertValue(object value, Type targetType)
        {
            if (value == null || value == DBNull.Value || string.IsNullOrEmpty(value.ToString()))
            {
                return targetType == typeof(string) ? string.Empty : DBNull.Value;
            }

            try
            {
                var stringValue = value.ToString();

                if (targetType == typeof(string)) return stringValue;
                if (targetType == typeof(int) && int.TryParse(stringValue, out int intValue)) return intValue;
                if (targetType == typeof(decimal) && decimal.TryParse(stringValue, out decimal decimalValue)) return decimalValue;
                if (targetType == typeof(DateTime) && DateTime.TryParse(stringValue, out DateTime dateValue)) return dateValue;
                if (targetType == typeof(bool))
                {
                    if (bool.TryParse(stringValue, out bool boolValue)) return boolValue;
                    return stringValue == "1";
                }
            }
            catch
            {
                // خطا در تبدیل
            }

            return targetType == typeof(string) ? value.ToString() : DBNull.Value;
        }

        private async Task CreateOrOverwriteTable(DataTable dataTable, string tableName, bool overwrite)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            // حذف جدول اگر وجود دارد و overwrite=true
            if (overwrite)
            {
                var dropSql = $@"
                IF OBJECT_ID('{tableName}', 'U') IS NOT NULL
                    DROP TABLE {tableName};";

                using var dropCommand = new SqlCommand(dropSql, connection);
                await dropCommand.ExecuteNonQueryAsync();
            }

            // ایجاد جدول جدید
            var createTableSql = GenerateCreateTableScript(dataTable, tableName);
            using var createCommand = new SqlCommand(createTableSql, connection);
            await createCommand.ExecuteNonQueryAsync();
        }

        private string GenerateCreateTableScript(DataTable dataTable, string tableName)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"CREATE TABLE [{tableName}] (");

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                var column = dataTable.Columns[i];
                var sqlType = GetSqlType(column.DataType);

                sb.Append($"    [{column.ColumnName}] {sqlType}");

                if (i < dataTable.Columns.Count - 1)
                    sb.AppendLine(",");
                else
                    sb.AppendLine();
            }

            sb.AppendLine(");");
            return sb.ToString();
        }

        private string GetSqlType(Type dataType)
        {
            if (dataType == typeof(int))
                return "INT";
            else if (dataType == typeof(double))
                return "float";
            else if (dataType == typeof(DateTime))
                return "DATETIME2";
            else if (dataType == typeof(bool))
                return "BIT";
            else
                return "NVARCHAR(MAX)";
        }

        private async Task BulkInsertToSqlServer(DataTable dataTable, string tableName)
        {
            if (dataTable.Rows.Count == 0)
                return;

            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            using var bulkCopy = new SqlBulkCopy(connection)
            {
                DestinationTableName = tableName,
                //BatchSize = 500000,
                BulkCopyTimeout = 0,
                EnableStreaming = true
            };

            // نگاشت خودکار ستون‌ها
            foreach (DataColumn column in dataTable.Columns)
            {
                bulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName);
            }

            await bulkCopy.WriteToServerAsync(dataTable);
        }

        private string CleanTableName(string tableName)
        {
            // حذف کاراکترهای غیرمجاز در نام جدول
            var cleaned = Regex.Replace(tableName, @"[^\w]", "_");

            // اطمینان از شروع نام با حرف
            if (char.IsDigit(cleaned[0]))
                cleaned = "T_" + cleaned;

            return cleaned;
        }

        private string CleanColumnName(string columnName, bool useBrackets = false)
        {
            if (string.IsNullOrEmpty(columnName))
                return useBrackets ? "[Column_Unknown]" : "Column_Unknown";

            // حذف فضاهای اضافه
            var cleaned = columnName.Trim();

            // لیست کلمات کلیدی SQL که نباید استفاده شوند
            var sqlKeywords = new HashSet<string>
    {
        "SELECT", "FROM", "WHERE", "INSERT", "UPDATE", "DELETE", "DROP",
        "CREATE", "TABLE", "COLUMN", "DATABASE", "INDEX", "VIEW", "PROCEDURE",
        "FUNCTION", "TRIGGER", "UNION", "JOIN", "GROUP", "ORDER", "BY", "HAVING",
        "DISTINCT", "TOP", "COUNT", "SUM", "AVG", "MIN", "MAX", "NULL", "NOT",
        "AND", "OR", "LIKE", "IN", "BETWEEN", "EXISTS", "CASE", "WHEN", "THEN",
        "END", "AS", "ON", "USING", "WITH", "RECURSIVE", "OVER", "PARTITION"
    };

            // اگر نام ستون یک کلمه کلیدی SQL است، پیشوند اضافه کن
            if (sqlKeywords.Contains(cleaned.ToUpper()))
            {
                cleaned = "Col_" + cleaned;
            }

            // جایگزینی کاراکترهای غیرمجاز با underscore
            // اجازه دادن به حروف فارسی/عربی، انگلیسی، اعداد و underscore
            cleaned = Regex.Replace(cleaned, @"[^\w\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]", "_");

            // حذف underscoreهای تکراری
            cleaned = Regex.Replace(cleaned, @"_+", "_");

            // حذف underscore از ابتدا و انتها
            cleaned = cleaned.Trim('_');

            // اگر بعد از پاکسازی خالی شد
            if (string.IsNullOrEmpty(cleaned))
                return useBrackets ? "[Column_Empty]" : "Column_Empty";

            // اطمینان از شروع نام با حرف یا underscore
            if (char.IsDigit(cleaned[0]))
                cleaned = "C_" + cleaned;

            // اگر طول نام بیش از 128 کاراکتر باشد
            if (cleaned.Length > 128)
                cleaned = cleaned.Substring(0, 128);

            // اضافه کردن براکت اگر درخواست شده
            if (useBrackets)
            {
                // اگر نام با براکت شروع نشده باشد، اضافه کن
                if (!cleaned.StartsWith("[") && !cleaned.EndsWith("]"))
                    cleaned = $"[{cleaned}]";
            }

            return cleaned;
        }

        public void TestConnection()
        {
            using var connection = new SqlConnection(_connectionString);
            connection.Open();
            connection.Close();
        }
        public async Task ProcessFilesWithAppendLogic(List<string> filePaths, bool detectColumnTypes = true)
        {
            // طبقه‌بندی فایل‌ها
            var categorizedFiles = CategorizeFiles(filePaths);

            // اول پردازش فایل‌های اصلی (بدون - و ـ)
            foreach (var mainFile in categorizedFiles.MainFiles)
            {
                var tableName = GetTableNameFromFileName(mainFile);
                await ProcessExcelFileWithTableCreation(mainFile, tableName, detectColumnTypes, true);
            }

            // سپس پردازش فایل‌های الحاقی (با - یا ـ)
            foreach (var appendFile in categorizedFiles.AppendFiles)
            {
                var mainTableName = GetMainTableNameFromAppendFile(appendFile);
                await AppendToExistingTable(appendFile, mainTableName, detectColumnTypes);
            }
        }

        private (List<string> MainFiles, List<string> AppendFiles) CategorizeFiles(List<string> filePaths)
        {
            var mainFiles = new List<string>();
            var appendFiles = new List<string>();

            foreach (var filePath in filePaths)
            {
                var fileName = Path.GetFileNameWithoutExtension(filePath);

                if (fileName.Contains("-") || fileName.Contains("ـ"))
                {
                    appendFiles.Add(filePath);
                }
                else
                {
                    mainFiles.Add(filePath);
                }
            }

            return (mainFiles, appendFiles);
        }

        private string GetMainTableNameFromAppendFile(string appendFilePath)
        {
            var fileName = Path.GetFileNameWithoutExtension(appendFilePath);

            // جدا کردن بخش اصلی نام فایل (قبل از - یا ـ)
            var separatorIndex = fileName.IndexOfAny(new[] { '-', 'ـ' });
            if (separatorIndex > 0)
            {
                return CleanTableName(fileName.Substring(0, separatorIndex));
            }

            return CleanTableName(fileName);
        }

        public async Task AppendToExistingTable(string filePath, string tableName, bool detectColumnTypes = true)
        {
            // خواندن فایل Excel
            var dataTable = ReadExcelToDataTable(filePath);

            // تشخیص و تبدیل انواع ستون‌ها
            if (detectColumnTypes)
            {
                dataTable = DetectAndConvertColumnTypes(dataTable);
            }

            // درج داده‌ها در جدول موجود
            await BulkInsertToSqlServer(dataTable, tableName);
        }

        // متد کمکی برای بررسی وجود جدول
        public async Task<bool> TableExists(string tableName)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            var sql = $@"
            SELECT COUNT(*) 
            FROM INFORMATION_SCHEMA.TABLES 
            WHERE TABLE_NAME = N'{tableName}'";

            using var command = new SqlCommand(sql, connection);
            var result = await command.ExecuteScalarAsync();

            return Convert.ToInt32(result) > 0;
        }
        private string GetTableNameFromFileName(string filePath)
        {
            var fileName = Path.GetFileNameWithoutExtension(filePath);
            return CleanTableName(fileName);
        }
    }
}
