namespace ExcelToSql
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            panelHeader = new Panel();
            lblSubtitle = new Label();
            panelMain = new Panel();
            groupBoxFiles = new GroupBox();
            listBoxFiles = new ListBox();
            panelFileButtons = new Panel();
            btnClearAll = new Button();
            btnRemove = new Button();
            btnBrowse = new Button();
            groupBoxSettings = new GroupBox();
            tableLayoutSettings = new TableLayoutPanel();
            chkUseAppendLogic = new CheckBox();
            chkParallelProcessing = new CheckBox();
            chkDetectTypes = new CheckBox();
            chkOverwriteTable = new CheckBox();
            lblNamingStrategy = new Label();
            cmbNamingStrategy = new ComboBox();
            lblParallelism = new Label();
            numParallelism = new NumericUpDown();
            panelActions = new Panel();
            btnTestConnection = new Button();
            btnImport = new Button();
            panelProgress = new Panel();
            lblFileCount = new Label();
            progressBar = new ProgressBar();
            lblStatus = new Label();
            panelHeader.SuspendLayout();
            panelMain.SuspendLayout();
            groupBoxFiles.SuspendLayout();
            panelFileButtons.SuspendLayout();
            groupBoxSettings.SuspendLayout();
            tableLayoutSettings.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)numParallelism).BeginInit();
            panelActions.SuspendLayout();
            panelProgress.SuspendLayout();
            SuspendLayout();
            // 
            // panelHeader
            // 
            panelHeader.BackColor = Color.FromArgb(0, 123, 255);
            panelHeader.Controls.Add(lblSubtitle);
            panelHeader.Dock = DockStyle.Top;
            panelHeader.Location = new Point(0, 0);
            panelHeader.Name = "panelHeader";
            panelHeader.Size = new Size(900, 55);
            panelHeader.TabIndex = 0;
            // 
            // lblSubtitle
            // 
            lblSubtitle.AutoSize = true;
            lblSubtitle.Font = new Font("Segoe UI", 15F, FontStyle.Regular, GraphicsUnit.Point, 0);
            lblSubtitle.ForeColor = Color.FromArgb(200, 200, 255);
            lblSubtitle.Location = new Point(478, 9);
            lblSubtitle.Name = "lblSubtitle";
            lblSubtitle.Size = new Size(309, 28);
            lblSubtitle.TabIndex = 1;
            lblSubtitle.Text = "انتقال داده‌های اکسل به SQL Server ";
            // 
            // panelMain
            // 
            panelMain.Controls.Add(groupBoxFiles);
            panelMain.Controls.Add(groupBoxSettings);
            panelMain.Dock = DockStyle.Fill;
            panelMain.Location = new Point(0, 55);
            panelMain.Name = "panelMain";
            panelMain.Padding = new Padding(20);
            panelMain.Size = new Size(900, 395);
            panelMain.TabIndex = 1;
            // 
            // groupBoxFiles
            // 
            groupBoxFiles.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            groupBoxFiles.Controls.Add(listBoxFiles);
            groupBoxFiles.Controls.Add(panelFileButtons);
            groupBoxFiles.Location = new Point(20, 20);
            groupBoxFiles.Name = "groupBoxFiles";
            groupBoxFiles.Size = new Size(500, 355);
            groupBoxFiles.TabIndex = 0;
            groupBoxFiles.TabStop = false;
            groupBoxFiles.Text = "فایل‌های انتخاب شده";
            // 
            // listBoxFiles
            // 
            listBoxFiles.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            listBoxFiles.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            listBoxFiles.FormattingEnabled = true;
            listBoxFiles.HorizontalScrollbar = true;
            listBoxFiles.ItemHeight = 15;
            listBoxFiles.Location = new Point(15, 60);
            listBoxFiles.Name = "listBoxFiles";
            listBoxFiles.Size = new Size(470, 259);
            listBoxFiles.TabIndex = 0;
            listBoxFiles.SelectedIndexChanged += listBoxFiles_SelectedIndexChanged;
            // 
            // panelFileButtons
            // 
            panelFileButtons.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            panelFileButtons.Controls.Add(btnClearAll);
            panelFileButtons.Controls.Add(btnRemove);
            panelFileButtons.Controls.Add(btnBrowse);
            panelFileButtons.Location = new Point(15, 25);
            panelFileButtons.Name = "panelFileButtons";
            panelFileButtons.Size = new Size(470, 35);
            panelFileButtons.TabIndex = 1;
            // 
            // btnClearAll
            // 
            btnClearAll.BackColor = Color.FromArgb(220, 53, 69);
            btnClearAll.FlatStyle = FlatStyle.Flat;
            btnClearAll.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            btnClearAll.ForeColor = Color.White;
            btnClearAll.Location = new Point(20, 5);
            btnClearAll.Name = "btnClearAll";
            btnClearAll.Size = new Size(140, 25);
            btnClearAll.TabIndex = 2;
            btnClearAll.Text = "حذف همه";
            btnClearAll.UseVisualStyleBackColor = false;
            btnClearAll.Click += btnClearAll_Click;
            // 
            // btnRemove
            // 
            btnRemove.BackColor = Color.FromArgb(108, 117, 125);
            btnRemove.FlatStyle = FlatStyle.Flat;
            btnRemove.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            btnRemove.ForeColor = Color.White;
            btnRemove.Location = new Point(170, 5);
            btnRemove.Name = "btnRemove";
            btnRemove.Size = new Size(140, 25);
            btnRemove.TabIndex = 1;
            btnRemove.Text = "حذف انتخاب شده";
            btnRemove.UseVisualStyleBackColor = false;
            btnRemove.Click += btnRemove_Click;
            // 
            // btnBrowse
            // 
            btnBrowse.BackColor = Color.FromArgb(0, 123, 255);
            btnBrowse.FlatStyle = FlatStyle.Flat;
            btnBrowse.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnBrowse.ForeColor = Color.White;
            btnBrowse.Location = new Point(320, 5);
            btnBrowse.Name = "btnBrowse";
            btnBrowse.Size = new Size(140, 25);
            btnBrowse.TabIndex = 0;
            btnBrowse.Text = "انتخاب فایل‌ها";
            btnBrowse.UseVisualStyleBackColor = false;
            btnBrowse.Click += btnBrowse_Click;
            // 
            // groupBoxSettings
            // 
            groupBoxSettings.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
            groupBoxSettings.Controls.Add(tableLayoutSettings);
            groupBoxSettings.Location = new Point(540, 20);
            groupBoxSettings.Name = "groupBoxSettings";
            groupBoxSettings.Size = new Size(340, 355);
            groupBoxSettings.TabIndex = 1;
            groupBoxSettings.TabStop = false;
            groupBoxSettings.Text = "تنظیمات پیشرفته";
            // 
            // tableLayoutSettings
            // 
            tableLayoutSettings.ColumnCount = 2;
            tableLayoutSettings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 39.5209579F));
            tableLayoutSettings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60.4790421F));
            tableLayoutSettings.Controls.Add(chkUseAppendLogic, 0, 0);
            tableLayoutSettings.Controls.Add(chkParallelProcessing, 0, 1);
            tableLayoutSettings.Controls.Add(chkDetectTypes, 0, 2);
            tableLayoutSettings.Controls.Add(chkOverwriteTable, 0, 3);
            tableLayoutSettings.Controls.Add(lblNamingStrategy, 0, 4);
            tableLayoutSettings.Controls.Add(cmbNamingStrategy, 0, 5);
            tableLayoutSettings.Controls.Add(lblParallelism, 0, 6);
            tableLayoutSettings.Controls.Add(numParallelism, 1, 6);
            tableLayoutSettings.Dock = DockStyle.Fill;
            tableLayoutSettings.Location = new Point(3, 19);
            tableLayoutSettings.Name = "tableLayoutSettings";
            tableLayoutSettings.RowCount = 8;
            tableLayoutSettings.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            tableLayoutSettings.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            tableLayoutSettings.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            tableLayoutSettings.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            tableLayoutSettings.RowStyles.Add(new RowStyle(SizeType.Absolute, 25F));
            tableLayoutSettings.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            tableLayoutSettings.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            tableLayoutSettings.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutSettings.Size = new Size(334, 333);
            tableLayoutSettings.TabIndex = 0;
            // 
            // chkUseAppendLogic
            // 
            chkUseAppendLogic.AutoSize = true;
            chkUseAppendLogic.Checked = true;
            chkUseAppendLogic.CheckState = CheckState.Checked;
            tableLayoutSettings.SetColumnSpan(chkUseAppendLogic, 2);
            chkUseAppendLogic.Dock = DockStyle.Fill;
            chkUseAppendLogic.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            chkUseAppendLogic.Location = new Point(3, 3);
            chkUseAppendLogic.Name = "chkUseAppendLogic";
            chkUseAppendLogic.Size = new Size(328, 29);
            chkUseAppendLogic.TabIndex = 0;
            chkUseAppendLogic.Text = "استفاده از منطق الحاق فایل‌ها (فایل‌های - و ـ)";
            chkUseAppendLogic.UseVisualStyleBackColor = true;
            // 
            // chkParallelProcessing
            // 
            chkParallelProcessing.AutoSize = true;
            chkParallelProcessing.Checked = true;
            chkParallelProcessing.CheckState = CheckState.Checked;
            tableLayoutSettings.SetColumnSpan(chkParallelProcessing, 2);
            chkParallelProcessing.Dock = DockStyle.Fill;
            chkParallelProcessing.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            chkParallelProcessing.Location = new Point(3, 38);
            chkParallelProcessing.Name = "chkParallelProcessing";
            chkParallelProcessing.Size = new Size(328, 29);
            chkParallelProcessing.TabIndex = 1;
            chkParallelProcessing.Text = "پردازش موازی فایل‌ها";
            chkParallelProcessing.UseVisualStyleBackColor = true;
            chkParallelProcessing.CheckedChanged += chkParallelProcessing_CheckedChanged;
            // 
            // chkDetectTypes
            // 
            chkDetectTypes.AutoSize = true;
            chkDetectTypes.Checked = true;
            chkDetectTypes.CheckState = CheckState.Checked;
            tableLayoutSettings.SetColumnSpan(chkDetectTypes, 2);
            chkDetectTypes.Dock = DockStyle.Fill;
            chkDetectTypes.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            chkDetectTypes.Location = new Point(3, 73);
            chkDetectTypes.Name = "chkDetectTypes";
            chkDetectTypes.Size = new Size(328, 29);
            chkDetectTypes.TabIndex = 2;
            chkDetectTypes.Text = "تشخیص خودکار نوع ستون‌ها";
            chkDetectTypes.UseVisualStyleBackColor = true;
            // 
            // chkOverwriteTable
            // 
            chkOverwriteTable.AutoSize = true;
            chkOverwriteTable.Checked = true;
            chkOverwriteTable.CheckState = CheckState.Checked;
            tableLayoutSettings.SetColumnSpan(chkOverwriteTable, 2);
            chkOverwriteTable.Dock = DockStyle.Fill;
            chkOverwriteTable.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            chkOverwriteTable.Location = new Point(3, 108);
            chkOverwriteTable.Name = "chkOverwriteTable";
            chkOverwriteTable.Size = new Size(328, 29);
            chkOverwriteTable.TabIndex = 3;
            chkOverwriteTable.Text = "بازنویسی جدول در صورت وجود";
            chkOverwriteTable.UseVisualStyleBackColor = true;
            // 
            // lblNamingStrategy
            // 
            lblNamingStrategy.AutoSize = true;
            tableLayoutSettings.SetColumnSpan(lblNamingStrategy, 2);
            lblNamingStrategy.Dock = DockStyle.Fill;
            lblNamingStrategy.Font = new Font("Segoe UI", 8F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblNamingStrategy.ForeColor = Color.FromArgb(64, 64, 64);
            lblNamingStrategy.Location = new Point(3, 140);
            lblNamingStrategy.Name = "lblNamingStrategy";
            lblNamingStrategy.Size = new Size(328, 25);
            lblNamingStrategy.TabIndex = 4;
            lblNamingStrategy.Text = "استراتژی نامگذاری ستون‌ها:";
            lblNamingStrategy.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // cmbNamingStrategy
            // 
            tableLayoutSettings.SetColumnSpan(cmbNamingStrategy, 2);
            cmbNamingStrategy.Dock = DockStyle.Fill;
            cmbNamingStrategy.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbNamingStrategy.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            cmbNamingStrategy.FormattingEnabled = true;
            cmbNamingStrategy.Items.AddRange(new object[] { "نام اصلی ستون‌ها", "نام پاکسازی شده", "نام استاندارد با براکت" });
            cmbNamingStrategy.Location = new Point(3, 168);
            cmbNamingStrategy.Name = "cmbNamingStrategy";
            cmbNamingStrategy.Size = new Size(328, 23);
            cmbNamingStrategy.TabIndex = 5;
            cmbNamingStrategy.SelectedIndexChanged += cmbNamingStrategy_SelectedIndexChanged;
            // 
            // lblParallelism
            // 
            lblParallelism.AutoSize = true;
            lblParallelism.Dock = DockStyle.Fill;
            lblParallelism.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            lblParallelism.Location = new Point(205, 200);
            lblParallelism.Name = "lblParallelism";
            lblParallelism.Size = new Size(126, 35);
            lblParallelism.TabIndex = 6;
            lblParallelism.Text = "تعداد فایل‌های همزمان:";
            lblParallelism.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // numParallelism
            // 
            numParallelism.Dock = DockStyle.Left;
            numParallelism.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            numParallelism.Location = new Point(30, 203);
            numParallelism.Maximum = new decimal(new int[] { 10, 0, 0, 0 });
            numParallelism.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            numParallelism.Name = "numParallelism";
            numParallelism.Size = new Size(169, 23);
            numParallelism.TabIndex = 7;
            numParallelism.Value = new decimal(new int[] { 5, 0, 0, 0 });
            numParallelism.ValueChanged += numParallelism_ValueChanged;
            // 
            // panelActions
            // 
            panelActions.BackColor = Color.FromArgb(248, 249, 250);
            panelActions.Controls.Add(btnTestConnection);
            panelActions.Controls.Add(btnImport);
            panelActions.Dock = DockStyle.Bottom;
            panelActions.Location = new Point(0, 500);
            panelActions.Name = "panelActions";
            panelActions.Size = new Size(900, 60);
            panelActions.TabIndex = 2;
            // 
            // btnTestConnection
            // 
            btnTestConnection.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnTestConnection.BackColor = Color.FromArgb(108, 117, 125);
            btnTestConnection.FlatStyle = FlatStyle.Flat;
            btnTestConnection.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            btnTestConnection.ForeColor = Color.White;
            btnTestConnection.Location = new Point(20, 15);
            btnTestConnection.Name = "btnTestConnection";
            btnTestConnection.Size = new Size(150, 35);
            btnTestConnection.TabIndex = 1;
            btnTestConnection.Text = "تست اتصال دیتابیس";
            btnTestConnection.UseVisualStyleBackColor = false;
            btnTestConnection.Click += btnTestConnection_Click;
            // 
            // btnImport
            // 
            btnImport.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnImport.BackColor = Color.FromArgb(40, 167, 69);
            btnImport.FlatStyle = FlatStyle.Flat;
            btnImport.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnImport.ForeColor = Color.White;
            btnImport.Location = new Point(680, 15);
            btnImport.Name = "btnImport";
            btnImport.Size = new Size(200, 35);
            btnImport.TabIndex = 0;
            btnImport.Text = "شروع پردازش فایل‌ها";
            btnImport.UseVisualStyleBackColor = false;
            btnImport.Click += btnImport_Click;
            // 
            // panelProgress
            // 
            panelProgress.BackColor = Color.White;
            panelProgress.BorderStyle = BorderStyle.FixedSingle;
            panelProgress.Controls.Add(lblFileCount);
            panelProgress.Controls.Add(progressBar);
            panelProgress.Controls.Add(lblStatus);
            panelProgress.Dock = DockStyle.Bottom;
            panelProgress.Location = new Point(0, 450);
            panelProgress.Name = "panelProgress";
            panelProgress.Size = new Size(900, 50);
            panelProgress.TabIndex = 3;
            // 
            // lblFileCount
            // 
            lblFileCount.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            lblFileCount.AutoSize = true;
            lblFileCount.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblFileCount.ForeColor = Color.FromArgb(0, 123, 255);
            lblFileCount.Location = new Point(750, 15);
            lblFileCount.Name = "lblFileCount";
            lblFileCount.Size = new Size(105, 15);
            lblFileCount.TabIndex = 2;
            lblFileCount.Text = "تعداد فایل‌ها: 0 مورد";
            // 
            // progressBar
            // 
            progressBar.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            progressBar.Location = new Point(120, 12);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(600, 20);
            progressBar.TabIndex = 1;
            progressBar.Visible = false;
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            lblStatus.Location = new Point(20, 15);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(81, 15);
            lblStatus.TabIndex = 0;
            lblStatus.Text = "آماده برای کار...";
            // 
            // Form1
            // 
            ClientSize = new Size(900, 560);
            Controls.Add(panelMain);
            Controls.Add(panelProgress);
            Controls.Add(panelActions);
            Controls.Add(panelHeader);
            Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            MinimumSize = new Size(916, 599);
            Name = "Form1";
            RightToLeft = RightToLeft.Yes;
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Excel to SQL Server Importer";
            Load += Form1_Load;
            panelHeader.ResumeLayout(false);
            panelHeader.PerformLayout();
            panelMain.ResumeLayout(false);
            groupBoxFiles.ResumeLayout(false);
            panelFileButtons.ResumeLayout(false);
            groupBoxSettings.ResumeLayout(false);
            tableLayoutSettings.ResumeLayout(false);
            tableLayoutSettings.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)numParallelism).EndInit();
            panelActions.ResumeLayout(false);
            panelProgress.ResumeLayout(false);
            panelProgress.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private Panel panelHeader;
        private Label lblSubtitle;
        private Panel panelMain;
        private GroupBox groupBoxFiles;
        private ListBox listBoxFiles;
        private Panel panelFileButtons;
        private Button btnClearAll;
        private Button btnRemove;
        private Button btnBrowse;
        private GroupBox groupBoxSettings;
        private TableLayoutPanel tableLayoutSettings;
        private CheckBox chkUseAppendLogic;
        private CheckBox chkParallelProcessing;
        private CheckBox chkDetectTypes;
        private CheckBox chkOverwriteTable;
        private Label lblNamingStrategy;
        private ComboBox cmbNamingStrategy;
        private Label lblParallelism;
        private NumericUpDown numParallelism;
        private Panel panelActions;
        private Button btnTestConnection;
        private Button btnImport;
        private Panel panelProgress;
        private Label lblFileCount;
        private ProgressBar progressBar;
        private Label lblStatus;
    }
}

