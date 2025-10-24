using ExcelToSqlBulkLibrary;
using ExcelToSqlByBulk_Console;
using Microsoft.Data.SqlClient;
using System;
using System.Threading.Tasks;

namespace ExcelToSqlConsole
{

    class Program
    {
        static async Task Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            try
            {
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                ShowHeader();

                // خواندن تنظیمات
                var config = ConfigManager.LoadConfig();
                ShowConfigSummary(config);

                // تأیید کاربر
                if (config.Verbose && !ConfirmExecution())
                    return;
              
                // پردازش فایل‌ها
                await ProcessFilesWithConfig(config);
                stopwatch.Stop();
                Console.WriteLine("✅ پردازش کامل شد!"+ stopwatch.Elapsed.TotalSeconds+" ثانیه");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ خطا: {ex.Message}");
                Environment.Exit(1);
            }
        }

        static void ShowHeader()
        {
            Console.WriteLine("==========================================");
            Console.WriteLine("    Excel to SQL Server Importer");
            Console.WriteLine("==========================================");
            Console.WriteLine();
        }

        static void ShowConfigSummary(ConfigModel config)
        {
            Console.WriteLine("📋 Configuration Summary:");
            Console.WriteLine($"   📁 Input Path: {config.InputPath}");
            Console.WriteLine($"   🔗 Database: {GetDatabaseName(config.ConnectionString)}");
            Console.WriteLine($"   📊 Table Name: {(string.IsNullOrEmpty(config.TableName) ? "Auto" : config.TableName)}");
            Console.WriteLine($"   🔄 Append Logic: {(config.UseAppendLogic ? "Enabled" : "Disabled")}");
            Console.WriteLine($"   🧠 Detect Column Types: {(config.DetectColumnTypes ? "Enabled" : "Disabled")}");
            Console.WriteLine($"   ⚡ Parallel Processing: {(config.ParallelProcessing ? "Enabled" : "Disabled")}");
            Console.WriteLine($"   🔢 Max Parallel Files: {config.MaxDegreeOfParallelism}");
            Console.WriteLine($"   🏷️ Naming Strategy: {config.NamingStrategy}");
            Console.WriteLine();
        }

        static bool ConfirmExecution()
        {
            Console.Write("❓ Do you want to continue? (y/n): ");
            var response = Console.ReadLine()?.ToLower().Trim();
            return response == "y" || response == "yes";
        }

        static async Task ProcessFilesWithConfig(ConfigModel config)
        {
            // ایجاد پردازشگر
            var processor = new ExcelToSqlBulkProcessor(config.ConnectionString);

            // دریافت فایل‌ها
            var files = GetExcelFiles(config.InputPath);

            if (files.Count == 0)
            {
                Console.WriteLine("⚠️ No Excel files found!");
                return;
            }

            Console.WriteLine($"🔍 Number of files found: {files.Count}");
            if (config.UseAppendLogic)
            {
                // پردازش با منطق الحاق
                await ProcessWithAppendLogic(processor, files, config);
            }
            else
            {
                // پردازش عادی
                await ProcessNormal(processor, files, config);
            }
        }

        static async Task ProcessWithAppendLogic(ExcelToSqlBulkProcessor processor,
            List<string> files, ConfigModel config)
        {
            // طبقه‌بندی فایل‌ها
            var categorizedFiles = CategorizeFiles(files);

            Console.WriteLine($"📋 Main Files: {categorizedFiles.MainFiles.Count}");
            Console.WriteLine($"📥 Append Files: {categorizedFiles.AppendFiles.Count}");
            // پردازش فایل‌های اصلی (با قابلیت موازی)
            if (categorizedFiles.MainFiles.Count > 0)
            {
                Console.WriteLine("🔄 Processing main files...");
                await ProcessFiles(processor, categorizedFiles.MainFiles, config, true);
            }

            // پردازش فایل‌های الحاقی (با قابلیت موازی)
            if (categorizedFiles.AppendFiles.Count > 0)
            {
                Console.WriteLine("🔄 Processing append files...");
                await ProcessAppendFilesParallel(processor, categorizedFiles.AppendFiles, config);
            }
        }

        static async Task ProcessAppendFilesParallel(ExcelToSqlBulkProcessor processor,
            List<string> appendFiles, ConfigModel config)
        {
            if (config.ParallelProcessing && appendFiles.Count > 1)
            {
                // پردازش موازی فایل‌های الحاقی
                await ProcessAppendFilesParallelInternal(processor, appendFiles, config);
            }
            else
            {
                // پردازش ترتیبی فایل‌های الحاقی
                await ProcessAppendFilesSequential(processor, appendFiles, config);
            }
        }

        static async Task ProcessAppendFilesParallelInternal(ExcelToSqlBulkProcessor processor,
            List<string> appendFiles, ConfigModel config)
        {
            Console.WriteLine($"⚡ Parallel processing append files - Max {config.MaxDegreeOfParallelism} files simultaneously");
            var semaphore = new System.Threading.SemaphoreSlim(config.MaxDegreeOfParallelism);
            var tasks = new List<Task>();

            foreach (var file in appendFiles)
            {
                await semaphore.WaitAsync();

                tasks.Add(Task.Run(async () =>
                {
                    try
                    {
                        await ProcessSingleAppendFile(processor, file, config);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error appending {Path.GetFileName(file)}:  {ex.Message}");
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                }));
            }

            await Task.WhenAll(tasks);
        }

        static async Task ProcessAppendFilesSequential(ExcelToSqlBulkProcessor processor,
            List<string> appendFiles, ConfigModel config)
        {
            Console.WriteLine("🔢 Sequential processing of append files");


            foreach (var file in appendFiles)
            {
                try
                {
                    await ProcessSingleAppendFile(processor, file, config);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error appending {Path.GetFileName(file)}: {ex.Message}");
                }
            }
        }

        static async Task ProcessSingleAppendFile(ExcelToSqlBulkProcessor processor,
            string file, ConfigModel config)
        {
            var mainTableName = GetMainTableNameFromAppendFile(file);
            Console.WriteLine($"📥 Appending: {Path.GetFileName(file)} → {mainTableName}");
            if (await processor.TableExists(mainTableName))
            {
                await processor.AppendToExistingTable(file, mainTableName, config.DetectColumnTypes);
                Console.WriteLine($"✅ Successfully appended: {Path.GetFileName(file)}");
            }
            else
            {
                Console.WriteLine($"⚠️ Creating new table: {mainTableName}");
                await processor.ProcessExcelFileWithTableCreation(
                    file, mainTableName, config.DetectColumnTypes, config.OverwriteTable, true);
                Console.WriteLine($"✅ Successfully created and appended: {Path.GetFileName(file)}");
            }
        }

        static async Task ProcessNormal(ExcelToSqlBulkProcessor processor,
            List<string> files, ConfigModel config)
        {
            Console.WriteLine("🔄 Start Proccesing...");
            await ProcessFiles(processor, files, config, false);
        }

        static async Task ProcessFiles(ExcelToSqlBulkProcessor processor,
            List<string> files, ConfigModel config, bool forceCreate)
        {
            if (config.ParallelProcessing && files.Count > 1)
            {
                await ProcessFilesParallel(processor, files, config, forceCreate);
            }
            else
            {
                await ProcessFilesSequential(processor, files, config, forceCreate);
            }
        }

        static async Task ProcessFilesParallel(ExcelToSqlBulkProcessor processor,
            List<string> files, ConfigModel config, bool forceCreate)
        {
            Console.WriteLine($"⚡ Paralell Process - حداکثر {config.MaxDegreeOfParallelism} فایل همزمان");

            var semaphore = new System.Threading.SemaphoreSlim(config.MaxDegreeOfParallelism);
            var tasks = new List<Task>();

            foreach (var file in files)
            {
                await semaphore.WaitAsync();

                tasks.Add(Task.Run(async () =>
                {
                    try
                    {
                        var fileName = Path.GetFileNameWithoutExtension(file);
                        Console.WriteLine($"📥 Main File:: {Path.GetFileName(file)} ");
                        var tableName = string.IsNullOrEmpty(config.TableName) ? fileName : config.TableName;

                        await processor.ProcessExcelFileWithTableCreation(
                            file, tableName, config.DetectColumnTypes, config.OverwriteTable, forceCreate);

                        Console.WriteLine($"✅ {Path.GetFileName(file)}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ {Path.GetFileName(file)} - {ex.Message}");
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                }));
            }

            await Task.WhenAll(tasks);
        }

        static async Task ProcessFilesSequential(ExcelToSqlBulkProcessor processor,
            List<string> files, ConfigModel config, bool forceCreate)
        {
            Console.WriteLine("🔢 ProcessFilesSequential");

            foreach (var file in files)
            {
                try
                {
                    var fileName = Path.GetFileNameWithoutExtension(file);
                    var tableName = string.IsNullOrEmpty(config.TableName) ? fileName : config.TableName;
                    Console.WriteLine($"📥 Main File:: {Path.GetFileName(file)} ");
                    await processor.ProcessExcelFileWithTableCreation(
                        file, tableName, config.DetectColumnTypes, config.OverwriteTable, forceCreate);

                    Console.WriteLine($"✅ {Path.GetFileName(file)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ {Path.GetFileName(file)} - {ex.Message}");
                }
            }
        }

        static List<string> GetExcelFiles(string path)
        {
            var files = new List<string>();

            if (File.Exists(path) && (path.EndsWith(".xlsx") || path.EndsWith(".xls")))
            {
                files.Add(path);
            }
            else if (Directory.Exists(path))
            {
                files.AddRange(Directory.GetFiles(path, "*.xlsx"));
                files.AddRange(Directory.GetFiles(path, "*.xls"));
            }
            else
            {
                throw new DirectoryNotFoundException($"DirectoryNotFoundException: {path}");
            }

            return files;
        }

        static (List<string> MainFiles, List<string> AppendFiles) CategorizeFiles(List<string> files)
        {
            var mainFiles = new List<string>();
            var appendFiles = new List<string>();

            foreach (var file in files)
            {
                var fileName = Path.GetFileNameWithoutExtension(file);

                if (fileName.Contains("-") || fileName.Contains("ـ"))
                {
                    appendFiles.Add(file);
                }
                else
                {
                    mainFiles.Add(file);
                }
            }

            return (mainFiles, appendFiles);
        }

        static string GetMainTableNameFromAppendFile(string filePath)
        {
            var fileName = Path.GetFileNameWithoutExtension(filePath);

            var separatorIndex = fileName.IndexOfAny(new[] { '-', 'ـ' });
            if (separatorIndex > 0)
            {
                return fileName.Substring(0, separatorIndex).Trim();
            }

            return fileName;
        }

        static string GetDatabaseName(string connectionString)
        {
            try
            {
                var builder = new SqlConnectionStringBuilder(connectionString);
                return builder.InitialCatalog;
            }
            catch
            {
                return "نامشخص";
            }
        }
    }
}