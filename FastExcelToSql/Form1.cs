using ExcelToSqlBulkLibrary;
using ExcelToSql;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static ExcelToSqlBulkLibrary.ExcelToSqlBulkProcessor;

namespace ExcelToSql
{
    public partial class Form1 : Form
    {
        private ExcelToSqlBulkLibrary.ExcelToSqlBulkProcessor _processor;
        private List<string> _selectedFiles = new List<string>();
        private ConfigModel _config;
        private ColumnNamingStrategy _currentNamingStrategy = ColumnNamingStrategy.Clean;


        public Form1()
        {
            InitializeComponent();
            LoadConfiguration();
            InitializeProcessor();
            ApplyModernStyling();
            UpdateUIFromConfig();
        }
        private void LoadConfiguration()
        {
            try
            {
                _config = ConfigManager.LoadConfig();
                if (_config == null)
                {
                    // اگر فایل تنظیمات وجود نداشت، برنامه بسته می‌شود
                    this.Close();
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در بارگذاری تنظیمات: {ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }
        private void UpdateUIFromConfig()
        {
            if (_config != null)
            {
                // پر کردن فیلدها با مقادیر پیش‌فرض از تنظیمات
                chkUseAppendLogic.Checked = _config.UseAppendLogic;
                chkParallelProcessing.Checked = _config.ParallelProcessing;
                chkDetectTypes.Checked = _config.DetectColumnTypes;
                chkOverwriteTable.Checked = _config.OverwriteTable;
                numParallelism.Value = _config.MaxDegreeOfParallelism;

                // تنظیم استراتژی نامگذاری
                cmbNamingStrategy.SelectedIndex = _config.NamingStrategy.ToLower() switch
                {
                    "original" => 0,
                    "clean" => 1,
                    "brackets" => 2,
                    _ => 1
                };

                // نمایش اطلاعات اتصال
                UpdateConnectionInfo();
            }
        }
        private void UpdateConnectionInfo()
        {
            if (_config != null)
            {
                var dbName = GetDatabaseNameFromConnectionString(_config.ConnectionString);
                // lblConnectionInfo.Text = $"اتصال به: {dbName}";
                // lblConnectionInfo.ForeColor = Color.Green;
            }
        }

        private string GetDatabaseNameFromConnectionString(string connectionString)
        {
            try
            {
                if (string.IsNullOrEmpty(connectionString))
                    return "نامشخص";

                var parts = connectionString.Split(';');
                foreach (var part in parts)
                {
                    var keyValue = part.Split('=');
                    if (keyValue.Length == 2)
                    {
                        var key = keyValue[0].Trim().ToLower();
                        var value = keyValue[1].Trim();

                        if (key == "database" || key == "initial catalog")
                            return value;
                    }
                }

                return "پیش‌فرض";
            }
            catch
            {
                return "نامشخص";
            }
        }

        private void ApplyModernStyling()
        {
            // استایل کلی فرم
            this.BackColor = Color.White;
            this.Font = new Font("Segoe UI", 9);
            this.Padding = new Padding(20);

            // استایل گروه‌بندی‌ها
            foreach (Control control in this.Controls)
            {
                if (control is GroupBox groupBox)
                {
                    groupBox.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    groupBox.ForeColor = Color.FromArgb(64, 64, 64);
                }
            }
        }
        private void InitializeProcessor()
        {
            if (_config != null)
            {
                _processor = new ExcelToSqlBulkProcessor(_config.ConnectionString)
                {

                };
            }
            // var connectionString = "Server=.;Database=Test;Integrated Security=true;TrustServerCertificate=true;";
            // _processor = new ExcelToSqlBulkLibrary.ExcelToSqlBulkProcessor(connectionString, _currentNamingStrategy);
        }

        private void cmbNamingStrategy_SelectedIndexChanged(object sender, EventArgs e)
        {
            _currentNamingStrategy = cmbNamingStrategy.SelectedIndex switch
            {
                0 => ColumnNamingStrategy.Original,
                1 => ColumnNamingStrategy.Clean,
                2 => ColumnNamingStrategy.Brackets,
                _ => ColumnNamingStrategy.Clean
            };

            // به روزرسانی پردازشگر با استراتژی جدید
            InitializeProcessor();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                openFileDialog.Multiselect = true;
                openFileDialog.Title = "Select Excel Files";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _selectedFiles.Clear();
                    _selectedFiles.AddRange(openFileDialog.FileNames);

                    // نمایش فایل‌های انتخاب شده
                    listBoxFiles.Items.Clear();
                    foreach (var file in _selectedFiles)
                    {
                        listBoxFiles.Items.Add(Path.GetFileName(file));
                    }

                    UpdateUI();
                }
            }
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (listBoxFiles.SelectedIndex != -1)
            {
                _selectedFiles.RemoveAt(listBoxFiles.SelectedIndex);
                listBoxFiles.Items.RemoveAt(listBoxFiles.SelectedIndex);
                UpdateUI();
            }
        }

        private void btnClearAll_Click(object sender, EventArgs e)
        {
            _selectedFiles.Clear();
            listBoxFiles.Items.Clear();
            UpdateUI();
        }

        private async void btnImport_Click(object sender, EventArgs e)
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            if (_selectedFiles.Count == 0)
            {
                MessageBox.Show("لطفاً حداقل یک فایل اکسل انتخاب کنید.", "هشدار",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                SetControlsEnabled(false);
                progressBar.Visible = true;
                progressBar.Value = 0;
                progressBar.Maximum = _selectedFiles.Count;

                int successCount = 0;
                int errorCount = 0;

                // ✅ بررسی آیا از منطق الحاق استفاده کنیم یا نه
                if (chkUseAppendLogic.Checked)
                {
                    // استفاده از منطق الحاق + پردازش موازی
                    (successCount, errorCount) = await ProcessFilesWithAppendLogicParallel();
                }
                else
                {
                    // پردازش عادی (موازی یا ترتیبی)
                    (successCount, errorCount) = await ProcessFilesNormal();
                }

                // نمایش نتیجه نهایی
                lblStatus.Text = $"پردازش کامل شد! موفق: {successCount}, ناموفق: {errorCount}";
                stopwatch.Stop();
                if (errorCount == 0)
                {
                    
                    MessageBox.Show($"تمام {successCount} فایل با موفقیت پردازش شدند. {stopwatch.Elapsed.TotalSeconds}ثانیه",
                        "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                {
                    MessageBox.Show($"{successCount} فایل موفق، {errorCount} فایل ناموفق     {stopwatch.Elapsed.TotalSeconds}ثانیه",
                        "نتیجه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطای کلی: {ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogError($"خطای کلی: {ex.Message}");
            }
            finally
            {
                SetControlsEnabled(true);
                progressBar.Visible = false;
            }
        }
        private async Task<(int successCount, int errorCount)> ProcessFilesNormal()
        {
            int successCount = 0;
            int errorCount = 0;

            if (chkParallelProcessing.Checked && _selectedFiles.Count > 1)
            {
                var options = new ParallelOptions
                {
                    MaxDegreeOfParallelism = (int)numParallelism.Value
                };

                await Parallel.ForEachAsync(_selectedFiles, options, async (filePath, token) =>
                {
                    try
                    {
                        var fileName = Path.GetFileNameWithoutExtension(filePath);
                        await _processor.ProcessExcelFileWithTableCreation(
                            filePath, fileName, chkDetectTypes.Checked, chkOverwriteTable.Checked);

                        Interlocked.Increment(ref successCount);

                        // ✅ آپدیت thread-safe
                        SafeUpdateUI(() => UpdateFileListStatus(filePath, "✅"));
                    }
                    catch (Exception ex)
                    {
                        Interlocked.Increment(ref errorCount);

                        // ✅ آپدیت thread-safe
                        SafeUpdateUI(() =>
                        {
                            UpdateFileListStatus(filePath, $"❌ {ex.Message}");
                        });

                        LogError($"خطا در پردازش {Path.GetFileName(filePath)}: {ex.Message}");
                    }
                    finally
                    {
                        // ✅ آپدیت thread-safe
                        SafeUpdateUI(() =>
                        {
                            progressBar.Value = successCount + errorCount;
                            lblStatus.Text = $"پردازش شده: {successCount + errorCount} از {_selectedFiles.Count}";
                        });
                    }
                });
            }
            else
            {
                // پردازش ترتیبی (همان قبلی)
                for (int i = 0; i < _selectedFiles.Count; i++)
                {
                    var filePath = _selectedFiles[i];
                    var fileName = Path.GetFileNameWithoutExtension(filePath);

                    try
                    {
                        await _processor.ProcessExcelFileWithTableCreation(
                            filePath, fileName, chkDetectTypes.Checked, chkOverwriteTable.Checked);

                        successCount++;
                        UpdateFileListStatus(filePath, "✅");
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        UpdateFileListStatus(filePath, $"❌ {ex.Message}");
                        LogError($"خطا در پردازش {fileName}: {ex.Message}");
                    }

                    progressBar.Value = i + 1;
                    lblStatus.Text = $"پردازش شده: {i + 1} از {_selectedFiles.Count}";
                    Application.DoEvents();
                }
            }

            return (successCount, errorCount);
        }
        private async Task<(int successCount, int errorCount)> ProcessFilesWithAppendLogicParallel()
        {
            int successCount = 0;
            int errorCount = 0;

            // طبقه‌بندی فایل‌ها
            var categorizedFiles = CategorizeFiles(_selectedFiles);

            //MessageBox.Show(
            //    $"فایل‌های اصلی: {categorizedFiles.MainFiles.Count}\n" +
            //    $"فایل‌های الحاقی: {categorizedFiles.AppendFiles.Count}",
            //    "طبقه‌بندی فایل‌ها",
            //    MessageBoxButtons.OK, MessageBoxIcon.Information);

            // ۱. پردازش موازی فایل‌های اصلی (اولویت اول)
            var mainResults = await ProcessFilesParallelAsync(
                categorizedFiles.MainFiles,
                ProcessMainFileAsync,
                "فایل اصلی"
            );
            successCount += mainResults.successCount;
            errorCount += mainResults.errorCount;

            // ۲. پردازش موازی فایل‌های الحاقی (اولویت دوم)
            var appendResults = await ProcessFilesParallelAsync(
                categorizedFiles.AppendFiles,
                ProcessAppendFileAsync,
                "فایل الحاقی"
            );
            successCount += appendResults.successCount;
            errorCount += appendResults.errorCount;

            return (successCount, errorCount);
        }

        private async Task<(int successCount, int errorCount)> ProcessFilesParallelAsync(
      List<string> files,
      Func<string, Task<bool>> processFileFunc,
      string fileType)
        {
            if (files.Count == 0)
                return (0, 0);

            int successCount = 0;
            int errorCount = 0;

            if (chkParallelProcessing.Checked && files.Count > 1)
            {
                var options = new ParallelOptions
                {
                    MaxDegreeOfParallelism = (int)numParallelism.Value
                };

                await Parallel.ForEachAsync(files, options, async (filePath, token) =>
                {
                    try
                    {
                        var result = await processFileFunc(filePath);
                        if (result)
                            Interlocked.Increment(ref successCount);
                        else
                            Interlocked.Increment(ref errorCount);
                    }
                    catch (Exception ex)
                    {
                        Interlocked.Increment(ref errorCount);

                        // ✅ استفاده از Invoke برای آپدیت UI
                        this.Invoke(new Action(() =>
                        {
                            UpdateFileListStatus(filePath, $"❌ {fileType} - {ex.Message}");
                        }));

                        LogError($"خطا در پردازش {fileType} {Path.GetFileName(filePath)}: {ex.Message}");
                    }
                    finally
                    {
                        // ✅ استفاده از Invoke برای آپدیت progress bar
                        this.Invoke(new Action(() =>
                        {
                            progressBar.Value = successCount + errorCount;
                            lblStatus.Text = $"{fileType}: پردازش شده {successCount + errorCount} از {files.Count}";
                        }));
                    }
                });
            }
            else
            {
                // پردازش ترتیبی (همان قبلی)
                for (int i = 0; i < files.Count; i++)
                {
                    var filePath = files[i];
                    try
                    {
                        var result = await processFileFunc(filePath);
                        if (result)
                            successCount++;
                        else
                            errorCount++;
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        UpdateFileListStatus(filePath, $"❌ {fileType} - {ex.Message}");
                        LogError($"خطا در پردازش {fileType} {Path.GetFileName(filePath)}: {ex.Message}");
                    }

                    progressBar.Value = successCount + errorCount;
                    lblStatus.Text = $"{fileType}: پردازش شده {successCount + errorCount} از {files.Count}";
                    Application.DoEvents();
                }
            }

            return (successCount, errorCount);
        }

        private async Task<bool> ProcessMainFileAsync(string filePath)
        {
            try
            {
                var fileName = Path.GetFileNameWithoutExtension(filePath);

                // فایل‌های اصلی: همیشه جدول رو ایجاد یا بازنویسی می‌کنن
                await _processor.ProcessExcelFileWithTableCreation(
                    filePath, fileName, chkDetectTypes.Checked, chkOverwriteTable.Checked, true);

                // ✅ آپدیت thread-safe
                SafeUpdateUI(() => UpdateFileListStatus(filePath, "✅ اصلی"));
                return true;
            }
            catch (Exception ex)
            {
                SafeUpdateUI(() => UpdateFileListStatus(filePath, $"❌ اصلی - {ex.Message}"));
                throw;
            }
        }

        private async Task<bool> ProcessAppendFileAsync(string filePath)
        {
            try
            {
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var mainTableName = GetMainTableNameFromAppendFile(filePath);

                // بررسی وجود جدول اصلی
                if (await _processor.TableExists(mainTableName))
                {
                    // الحاق به جدول موجود
                    await _processor.AppendToExistingTable(filePath, mainTableName, chkDetectTypes.Checked);
                    SafeUpdateUI(() => UpdateFileListStatus(filePath, $"✅ الحاق به {mainTableName}"));
                }
                else
                {
                    // اگر جدول اصلی وجود ندارد، ایجاد جدید
                    await _processor.ProcessExcelFileWithTableCreation(
                        filePath, mainTableName, chkDetectTypes.Checked, false, true);
                    SafeUpdateUI(() => UpdateFileListStatus(filePath, $"✅ ایجاد {mainTableName} (اصلی یافت نشد)"));
                }
                return true;
            }
            catch (Exception ex)
            {
                SafeUpdateUI(() => UpdateFileListStatus(filePath, $"❌ الحاق"));
                throw new Exception($"خطا در الحاق: {ex.Message}", ex);
            }
        }
        private void SafeUpdateUI(Action action)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(action);
            }
            else
            {
                action();
            }
        }

        // استفاده:
        //SafeUpdateUI(() => {
        //    progressBar.Value = currentValue;
        //    lblStatus.Text = "در حال پردازش...";
        //});
        private (List<string> MainFiles, List<string> AppendFiles) CategorizeFiles(List<string> files)
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



        private string GetMainTableNameFromAppendFile(string filePath)
        {
            var fileName = Path.GetFileNameWithoutExtension(filePath);

            // جدا کردن بخش اصلی نام فایل (قبل از - یا ـ)
            var separatorIndex = fileName.IndexOfAny(new[] { '-', 'ـ' });
            if (separatorIndex > 0)
            {
                return fileName.Substring(0, separatorIndex).Trim();
            }

            return fileName;
        }

        private void UpdateFileListStatus(string filePath, string status)
        {
            var fileName = Path.GetFileName(filePath);
            var index = _selectedFiles.IndexOf(filePath);

            if (index >= 0)
            {
                // ✅ استفاده از Invoke برای آپدیت thread-safe
                if (listBoxFiles.InvokeRequired)
                {
                    listBoxFiles.Invoke(new Action(() =>
                    {
                        listBoxFiles.Items[index] = $"{status} {fileName}";
                    }));
                }
                else
                {
                    listBoxFiles.Items[index] = $"{status} {fileName}";
                }
            }
        }
        private void SetControlsEnabled(bool enabled)
        {
            btnBrowse.Enabled = enabled;
            btnImport.Enabled = enabled && _selectedFiles.Count > 0;
            btnRemove.Enabled = enabled;
            btnClearAll.Enabled = enabled;
            chkDetectTypes.Enabled = enabled;
            chkOverwriteTable.Enabled = enabled;
            chkUseAppendLogic.Enabled = enabled;
            chkParallelProcessing.Enabled = enabled;
        }

        private void UpdateUI()
        {
            btnImport.Enabled = _selectedFiles.Count > 0;
            btnRemove.Enabled = listBoxFiles.SelectedIndex != -1;
            btnClearAll.Enabled = _selectedFiles.Count > 0;

            lblFileCount.Text = $"تعداد فایل‌ها: {_selectedFiles.Count}";
        }

        private void listBoxFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnRemove.Enabled = listBoxFiles.SelectedIndex != -1;
        }

        private void LogError(string message)
        {
            // در اینجا می‌توانید خطاها را در فایل لاگ ذخیره کنید
            System.Diagnostics.Debug.WriteLine($"[ERROR] {DateTime.Now}: {message}");
        }
        private void chkParallelProcessing_CheckedChanged(object sender, EventArgs e)
        {
            numParallelism.Enabled = chkParallelProcessing.Checked;
            lblParallelism.Enabled = chkParallelProcessing.Checked;
        }
        private void btnTestConnection_Click(object sender, EventArgs e)
        {
            try
            {
                _processor.TestConnection();
                MessageBox.Show("اتصال به دیتابیس موفقیت‌آمیز بود.", "تست اتصال",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در اتصال به دیتابیس: {ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void numParallelism_ValueChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}