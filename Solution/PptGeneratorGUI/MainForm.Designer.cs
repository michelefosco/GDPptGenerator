namespace PptGeneratorGUI
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.toolStripProgressBar = new System.Windows.Forms.ToolStripProgressBar();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.txtStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.lblVersion = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolTipDefault = new System.Windows.Forms.ToolTip(this.components);
            this.btnSelectFileBudget = new System.Windows.Forms.Button();
            this.btnSelectDestinationFolder = new System.Windows.Forms.Button();
            this.lblCartellaOutputPath = new System.Windows.Forms.Label();
            this.btnSelectForecastFile = new System.Windows.Forms.Button();
            this.btnSelectFileSuperDettagli = new System.Windows.Forms.Button();
            this.btnSelectFileRanRate = new System.Windows.Forms.Button();
            this.lblDataPeriodo = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnOpenCalendar = new System.Windows.Forms.Button();
            this.btnOpenFileRanRate = new System.Windows.Forms.Button();
            this.btnOpenFileRanRateFolder = new System.Windows.Forms.Button();
            this.btnOpenFileSuperDettagli = new System.Windows.Forms.Button();
            this.btnOpenFileSuperDettagliFolder = new System.Windows.Forms.Button();
            this.btnOpenFileForecast = new System.Windows.Forms.Button();
            this.btnOpenFileForecastFolder = new System.Windows.Forms.Button();
            this.btnOpenFileBudget = new System.Windows.Forms.Button();
            this.btnOpenDestFolder = new System.Windows.Forms.Button();
            this.btnOpenFileBudgetFolder = new System.Windows.Forms.Button();
            this.cmbFileBudgetPath = new System.Windows.Forms.ComboBox();
            this.lblFileBudgetPath = new System.Windows.Forms.Label();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.label2 = new System.Windows.Forms.Label();
            this.menuStrip = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItemClear = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItemOpenConfigFolder = new System.Windows.Forms.ToolStripMenuItem();
            this.cmbDestinationFolderPath = new System.Windows.Forms.ComboBox();
            this.wbExecutionResult = new System.Windows.Forms.WebBrowser();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.btnCreatePresentationBackgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.btnCopyError = new System.Windows.Forms.LinkLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.cmbFileForecastPath = new System.Windows.Forms.ComboBox();
            this.lblFileForecastPath = new System.Windows.Forms.Label();
            this.cmbFileSuperDettagliPath = new System.Windows.Forms.ComboBox();
            this.lblFileSuperDettagliPath = new System.Windows.Forms.Label();
            this.cmbFileRanRatePath = new System.Windows.Forms.ComboBox();
            this.lblFileRanRatePath = new System.Windows.Forms.Label();
            this.btnCreaPresentazione = new System.Windows.Forms.Button();
            this.lblFiltriApplicabili = new System.Windows.Forms.Label();
            this.dgvFiltri = new System.Windows.Forms.DataGridView();
            this.Tabella = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Campo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OpenFiltersSelection = new System.Windows.Forms.DataGridViewButtonColumn();
            this.ValoriSelezionati = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnClear = new System.Windows.Forms.LinkLabel();
            this.calendarPeriodo = new System.Windows.Forms.MonthCalendar();
            this.pnlCalendar = new System.Windows.Forms.Panel();
            this.gbPaths = new System.Windows.Forms.GroupBox();
            this.gbOptions = new System.Windows.Forms.GroupBox();
            this.btnNext = new System.Windows.Forms.Button();
            this.btnNextBackgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.bfbDestFolder = new WK.Libraries.BetterFolderBrowserNS.BetterFolderBrowser(this.components);
            this.menuStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvFiltri)).BeginInit();
            this.pnlCalendar.SuspendLayout();
            this.gbPaths.SuspendLayout();
            this.gbOptions.SuspendLayout();
            this.SuspendLayout();
            // 
            // statusStrip
            // 
            this.statusStrip.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.statusStrip.Location = new System.Drawing.Point(0, 1013);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Padding = new System.Windows.Forms.Padding(2, 0, 14, 0);
            this.statusStrip.Size = new System.Drawing.Size(1602, 22);
            this.statusStrip.TabIndex = 21;
            this.statusStrip.Text = "statusStrip1";
            // 
            // toolStripProgressBar
            // 
            this.toolStripProgressBar.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.toolStripProgressBar.Name = "toolStripProgressBar";
            this.toolStripProgressBar.Size = new System.Drawing.Size(100, 16);
            this.toolStripProgressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.toolStripProgressBar.Visible = false;
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(43, 17);
            this.toolStripStatusLabel1.Text = "  Stato:";
            // 
            // txtStatusLabel
            // 
            this.txtStatusLabel.Name = "txtStatusLabel";
            this.txtStatusLabel.Size = new System.Drawing.Size(49, 17);
            this.txtStatusLabel.Text = "[STATO]";
            // 
            // lblVersion
            // 
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(876, 17);
            this.lblVersion.Spring = true;
            this.lblVersion.Text = "[VERSIONE]";
            this.lblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnSelectFileBudget
            // 
            this.btnSelectFileBudget.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFileBudget.Location = new System.Drawing.Point(1481, 14);
            this.btnSelectFileBudget.Name = "btnSelectFileBudget";
            this.btnSelectFileBudget.Size = new System.Drawing.Size(29, 23);
            this.btnSelectFileBudget.TabIndex = 1;
            this.btnSelectFileBudget.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectFileBudget, "Apre la finestra di selezione file");
            this.btnSelectFileBudget.UseVisualStyleBackColor = true;
            this.btnSelectFileBudget.Click += new System.EventHandler(this.btnSelectFileBudget_Click);
            // 
            // btnSelectDestinationFolder
            // 
            this.btnSelectDestinationFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectDestinationFolder.Location = new System.Drawing.Point(1481, 120);
            this.btnSelectDestinationFolder.Name = "btnSelectDestinationFolder";
            this.btnSelectDestinationFolder.Size = new System.Drawing.Size(29, 23);
            this.btnSelectDestinationFolder.TabIndex = 17;
            this.btnSelectDestinationFolder.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectDestinationFolder, "Apre la finestra di selezione cartella");
            this.btnSelectDestinationFolder.UseVisualStyleBackColor = true;
            this.btnSelectDestinationFolder.Click += new System.EventHandler(this.btnSelectDestinationFolder_Click);
            // 
            // lblCartellaOutputPath
            // 
            this.lblCartellaOutputPath.AutoSize = true;
            this.lblCartellaOutputPath.Location = new System.Drawing.Point(5, 125);
            this.lblCartellaOutputPath.Name = "lblCartellaOutputPath";
            this.lblCartellaOutputPath.Size = new System.Drawing.Size(71, 13);
            this.lblCartellaOutputPath.TabIndex = 11;
            this.lblCartellaOutputPath.Text = "Output folder:";
            this.toolTipDefault.SetToolTip(this.lblCartellaOutputPath, "Cartella nella quale vettanno salvati i file di output");
            // 
            // btnSelectForecastFile
            // 
            this.btnSelectForecastFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectForecastFile.Location = new System.Drawing.Point(1481, 40);
            this.btnSelectForecastFile.Name = "btnSelectForecastFile";
            this.btnSelectForecastFile.Size = new System.Drawing.Size(29, 23);
            this.btnSelectForecastFile.TabIndex = 5;
            this.btnSelectForecastFile.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectForecastFile, "Apre la finestra di selezione file");
            this.btnSelectForecastFile.UseVisualStyleBackColor = true;
            this.btnSelectForecastFile.Click += new System.EventHandler(this.btnSelectForecastFile_Click);
            // 
            // btnSelectFileSuperDettagli
            // 
            this.btnSelectFileSuperDettagli.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFileSuperDettagli.Location = new System.Drawing.Point(1481, 94);
            this.btnSelectFileSuperDettagli.Name = "btnSelectFileSuperDettagli";
            this.btnSelectFileSuperDettagli.Size = new System.Drawing.Size(29, 23);
            this.btnSelectFileSuperDettagli.TabIndex = 13;
            this.btnSelectFileSuperDettagli.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectFileSuperDettagli, "Apre la finestra di selezione file");
            this.btnSelectFileSuperDettagli.UseVisualStyleBackColor = true;
            this.btnSelectFileSuperDettagli.Click += new System.EventHandler(this.btnSelectFileSuperDettagli_Click);
            // 
            // btnSelectFileRanRate
            // 
            this.btnSelectFileRanRate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFileRanRate.Location = new System.Drawing.Point(1481, 67);
            this.btnSelectFileRanRate.Name = "btnSelectFileRanRate";
            this.btnSelectFileRanRate.Size = new System.Drawing.Size(29, 23);
            this.btnSelectFileRanRate.TabIndex = 9;
            this.btnSelectFileRanRate.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectFileRanRate, "Apre la finestra di selezione file");
            this.btnSelectFileRanRate.UseVisualStyleBackColor = true;
            this.btnSelectFileRanRate.Click += new System.EventHandler(this.btnSelectFileRanRate_Click);
            // 
            // lblDataPeriodo
            // 
            this.lblDataPeriodo.AutoSize = true;
            this.lblDataPeriodo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDataPeriodo.Location = new System.Drawing.Point(118, 23);
            this.lblDataPeriodo.Name = "lblDataPeriodo";
            this.lblDataPeriodo.Size = new System.Drawing.Size(108, 13);
            this.lblDataPeriodo.TabIndex = 46;
            this.lblDataPeriodo.Text = "[DATA PERIODO]";
            this.toolTipDefault.SetToolTip(this.lblDataPeriodo, "Cartella nella quale vettanno salvati i file di output");
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(14, 23);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(40, 13);
            this.label4.TabIndex = 47;
            this.label4.Text = "Period:";
            this.toolTipDefault.SetToolTip(this.label4, "Cartella nella quale vettanno salvati i file di output");
            // 
            // btnOpenCalendar
            // 
            this.btnOpenCalendar.Image = global::PptGeneratorGUI.Properties.Resources.CalendarIcon;
            this.btnOpenCalendar.Location = new System.Drawing.Point(232, 18);
            this.btnOpenCalendar.Name = "btnOpenCalendar";
            this.btnOpenCalendar.Size = new System.Drawing.Size(29, 23);
            this.btnOpenCalendar.TabIndex = 19;
            this.toolTipDefault.SetToolTip(this.btnOpenCalendar, "Apri il calendario per selezionare la data del periodo di riferimento");
            this.btnOpenCalendar.UseVisualStyleBackColor = true;
            this.btnOpenCalendar.Click += new System.EventHandler(this.btnOpenCalendar_Click);
            // 
            // btnOpenFileRanRate
            // 
            this.btnOpenFileRanRate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileRanRate.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileRanRate.Image")));
            this.btnOpenFileRanRate.Location = new System.Drawing.Point(1552, 66);
            this.btnOpenFileRanRate.Name = "btnOpenFileRanRate";
            this.btnOpenFileRanRate.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileRanRate.TabIndex = 11;
            this.toolTipDefault.SetToolTip(this.btnOpenFileRanRate, "Apre il file report");
            this.btnOpenFileRanRate.UseVisualStyleBackColor = true;
            this.btnOpenFileRanRate.Click += new System.EventHandler(this.btnOpenFileRanRate_Click);
            // 
            // btnOpenFileRanRateFolder
            // 
            this.btnOpenFileRanRateFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileRanRateFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileRanRateFolder.Image")));
            this.btnOpenFileRanRateFolder.Location = new System.Drawing.Point(1516, 66);
            this.btnOpenFileRanRateFolder.Name = "btnOpenFileRanRateFolder";
            this.btnOpenFileRanRateFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileRanRateFolder.TabIndex = 10;
            this.toolTipDefault.SetToolTip(this.btnOpenFileRanRateFolder, "Apre la cartella contenente il file report");
            this.btnOpenFileRanRateFolder.UseVisualStyleBackColor = true;
            this.btnOpenFileRanRateFolder.Click += new System.EventHandler(this.btnOpenFileRanRateFolder_Click);
            // 
            // btnOpenFileSuperDettagli
            // 
            this.btnOpenFileSuperDettagli.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileSuperDettagli.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileSuperDettagli.Image")));
            this.btnOpenFileSuperDettagli.Location = new System.Drawing.Point(1551, 93);
            this.btnOpenFileSuperDettagli.Name = "btnOpenFileSuperDettagli";
            this.btnOpenFileSuperDettagli.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileSuperDettagli.TabIndex = 15;
            this.toolTipDefault.SetToolTip(this.btnOpenFileSuperDettagli, "Apre il file report");
            this.btnOpenFileSuperDettagli.UseVisualStyleBackColor = true;
            this.btnOpenFileSuperDettagli.Click += new System.EventHandler(this.btnOpenFileSuperDettagli_Click);
            // 
            // btnOpenFileSuperDettagliFolder
            // 
            this.btnOpenFileSuperDettagliFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileSuperDettagliFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileSuperDettagliFolder.Image")));
            this.btnOpenFileSuperDettagliFolder.Location = new System.Drawing.Point(1516, 94);
            this.btnOpenFileSuperDettagliFolder.Name = "btnOpenFileSuperDettagliFolder";
            this.btnOpenFileSuperDettagliFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileSuperDettagliFolder.TabIndex = 14;
            this.toolTipDefault.SetToolTip(this.btnOpenFileSuperDettagliFolder, "Apre la cartella contenente il file report");
            this.btnOpenFileSuperDettagliFolder.UseVisualStyleBackColor = true;
            this.btnOpenFileSuperDettagliFolder.Click += new System.EventHandler(this.btnOpenFileSuperDettagliFolder_Click);
            // 
            // btnOpenFileForecast
            // 
            this.btnOpenFileForecast.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileForecast.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileForecast.Image")));
            this.btnOpenFileForecast.Location = new System.Drawing.Point(1552, 40);
            this.btnOpenFileForecast.Name = "btnOpenFileForecast";
            this.btnOpenFileForecast.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileForecast.TabIndex = 7;
            this.toolTipDefault.SetToolTip(this.btnOpenFileForecast, "Apre il file report");
            this.btnOpenFileForecast.UseVisualStyleBackColor = true;
            this.btnOpenFileForecast.Click += new System.EventHandler(this.btnOpenFileForecast_Click);
            // 
            // btnOpenFileForecastFolder
            // 
            this.btnOpenFileForecastFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileForecastFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileForecastFolder.Image")));
            this.btnOpenFileForecastFolder.Location = new System.Drawing.Point(1516, 40);
            this.btnOpenFileForecastFolder.Name = "btnOpenFileForecastFolder";
            this.btnOpenFileForecastFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileForecastFolder.TabIndex = 6;
            this.toolTipDefault.SetToolTip(this.btnOpenFileForecastFolder, "Apre la cartella contenente il file report");
            this.btnOpenFileForecastFolder.UseVisualStyleBackColor = true;
            this.btnOpenFileForecastFolder.Click += new System.EventHandler(this.btnOpenFileForecastFolder_Click);
            // 
            // btnOpenFileBudget
            // 
            this.btnOpenFileBudget.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileBudget.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileBudget.Image")));
            this.btnOpenFileBudget.Location = new System.Drawing.Point(1552, 14);
            this.btnOpenFileBudget.Name = "btnOpenFileBudget";
            this.btnOpenFileBudget.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileBudget.TabIndex = 3;
            this.toolTipDefault.SetToolTip(this.btnOpenFileBudget, "Apre il file report");
            this.btnOpenFileBudget.UseVisualStyleBackColor = true;
            this.btnOpenFileBudget.Click += new System.EventHandler(this.btnOpenFileBudget_Click);
            // 
            // btnOpenDestFolder
            // 
            this.btnOpenDestFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenDestFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenDestFolder.Image")));
            this.btnOpenDestFolder.Location = new System.Drawing.Point(1516, 121);
            this.btnOpenDestFolder.Name = "btnOpenDestFolder";
            this.btnOpenDestFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenDestFolder.TabIndex = 18;
            this.toolTipDefault.SetToolTip(this.btnOpenDestFolder, "Apre la cartella di destinazione");
            this.btnOpenDestFolder.UseVisualStyleBackColor = true;
            this.btnOpenDestFolder.Click += new System.EventHandler(this.btnOpenDestFolder_Click);
            // 
            // btnOpenFileBudgetFolder
            // 
            this.btnOpenFileBudgetFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileBudgetFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileBudgetFolder.Image")));
            this.btnOpenFileBudgetFolder.Location = new System.Drawing.Point(1516, 14);
            this.btnOpenFileBudgetFolder.Name = "btnOpenFileBudgetFolder";
            this.btnOpenFileBudgetFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileBudgetFolder.TabIndex = 2;
            this.toolTipDefault.SetToolTip(this.btnOpenFileBudgetFolder, "Apre la cartella contenente il file report");
            this.btnOpenFileBudgetFolder.UseVisualStyleBackColor = true;
            this.btnOpenFileBudgetFolder.Click += new System.EventHandler(this.btnOpenFileBudgetFolder_Click);
            // 
            // cmbFileBudgetPath
            // 
            this.cmbFileBudgetPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFileBudgetPath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbFileBudgetPath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbFileBudgetPath.FormattingEnabled = true;
            this.cmbFileBudgetPath.Location = new System.Drawing.Point(115, 16);
            this.cmbFileBudgetPath.Name = "cmbFileBudgetPath";
            this.cmbFileBudgetPath.Size = new System.Drawing.Size(1357, 21);
            this.cmbFileBudgetPath.TabIndex = 0;
            this.cmbFileBudgetPath.SelectedIndexChanged += new System.EventHandler(this.cmbFileBudgetPath_SelectedIndexChanged);
            this.cmbFileBudgetPath.TextUpdate += new System.EventHandler(this.cmbFileBudgetPath_TextUpdate);
            // 
            // lblFileBudgetPath
            // 
            this.lblFileBudgetPath.AutoSize = true;
            this.lblFileBudgetPath.Location = new System.Drawing.Point(6, 19);
            this.lblFileBudgetPath.Name = "lblFileBudgetPath";
            this.lblFileBudgetPath.Size = new System.Drawing.Size(70, 13);
            this.lblFileBudgetPath.TabIndex = 6;
            this.lblFileBudgetPath.Text = "\"Budget\" file:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 528);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(45, 13);
            this.label2.TabIndex = 20;
            this.label2.Text = "Results:";
            // 
            // menuStrip
            // 
            this.menuStrip.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.menuStrip.Dock = System.Windows.Forms.DockStyle.None;
            this.menuStrip.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1});
            this.menuStrip.Location = new System.Drawing.Point(0, 0);
            this.menuStrip.Name = "menuStrip";
            this.menuStrip.Padding = new System.Windows.Forms.Padding(4, 2, 0, 2);
            this.menuStrip.Size = new System.Drawing.Size(67, 24);
            this.menuStrip.TabIndex = 0;
            this.menuStrip.Text = "menuStrip1";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItemClear,
            this.toolStripMenuItemOpenConfigFolder});
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(61, 20);
            this.toolStripMenuItem1.Text = "&Settings";
            // 
            // toolStripMenuItemClear
            // 
            this.toolStripMenuItemClear.Name = "toolStripMenuItemClear";
            this.toolStripMenuItemClear.Size = new System.Drawing.Size(193, 22);
            this.toolStripMenuItemClear.Text = "&Clean path lists history";
            this.toolStripMenuItemClear.Click += new System.EventHandler(this.toolStripMenuItemClear_Click);
            // 
            // toolStripMenuItemOpenConfigFolder
            // 
            this.toolStripMenuItemOpenConfigFolder.Name = "toolStripMenuItemOpenConfigFolder";
            this.toolStripMenuItemOpenConfigFolder.Size = new System.Drawing.Size(193, 22);
            this.toolStripMenuItemOpenConfigFolder.Text = "Open &templates folder";
            this.toolStripMenuItemOpenConfigFolder.Click += new System.EventHandler(this.toolStripMenuItemOpenConfigFolder_Click);
            // 
            // cmbDestinationFolderPath
            // 
            this.cmbDestinationFolderPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbDestinationFolderPath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbDestinationFolderPath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbDestinationFolderPath.FormattingEnabled = true;
            this.cmbDestinationFolderPath.Location = new System.Drawing.Point(115, 122);
            this.cmbDestinationFolderPath.Name = "cmbDestinationFolderPath";
            this.cmbDestinationFolderPath.Size = new System.Drawing.Size(1357, 21);
            this.cmbDestinationFolderPath.TabIndex = 16;
            this.cmbDestinationFolderPath.SelectedIndexChanged += new System.EventHandler(this.cmbDestinationFolderPath_SelectedIndexChanged);
            this.cmbDestinationFolderPath.TextUpdate += new System.EventHandler(this.cmbDestinationFolderPath_TextUpdate);
            // 
            // wbExecutionResult
            // 
            this.wbExecutionResult.AllowWebBrowserDrop = false;
            this.wbExecutionResult.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.wbExecutionResult.IsWebBrowserContextMenuEnabled = false;
            this.wbExecutionResult.Location = new System.Drawing.Point(12, 547);
            this.wbExecutionResult.MinimumSize = new System.Drawing.Size(20, 20);
            this.wbExecutionResult.Name = "wbExecutionResult";
            this.wbExecutionResult.Size = new System.Drawing.Size(1578, 454);
            this.wbExecutionResult.TabIndex = 23;
            this.wbExecutionResult.Navigating += new System.Windows.Forms.WebBrowserNavigatingEventHandler(this.wbExecutionResult_Navigating_1);
            // 
            // btnCopyError
            // 
            this.btnCopyError.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCopyError.AutoSize = true;
            this.btnCopyError.Location = new System.Drawing.Point(1418, 988);
            this.btnCopyError.Name = "btnCopyError";
            this.btnCopyError.Size = new System.Drawing.Size(148, 13);
            this.btnCopyError.TabIndex = 25;
            this.btnCopyError.TabStop = true;
            this.btnCopyError.Text = "Copy the error in the clipboard";
            this.btnCopyError.Visible = false;
            this.btnCopyError.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.btnCopyError_LinkClicked);
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(0, 17);
            // 
            // cmbFileForecastPath
            // 
            this.cmbFileForecastPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFileForecastPath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbFileForecastPath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbFileForecastPath.FormattingEnabled = true;
            this.cmbFileForecastPath.Location = new System.Drawing.Point(115, 42);
            this.cmbFileForecastPath.Name = "cmbFileForecastPath";
            this.cmbFileForecastPath.Size = new System.Drawing.Size(1357, 21);
            this.cmbFileForecastPath.TabIndex = 4;
            this.cmbFileForecastPath.SelectedIndexChanged += new System.EventHandler(this.cmbFileForecastPath_SelectedIndexChanged);
            this.cmbFileForecastPath.TextUpdate += new System.EventHandler(this.cmbFileForecastPath_TextUpdate);
            // 
            // lblFileForecastPath
            // 
            this.lblFileForecastPath.AutoSize = true;
            this.lblFileForecastPath.Location = new System.Drawing.Point(6, 45);
            this.lblFileForecastPath.Name = "lblFileForecastPath";
            this.lblFileForecastPath.Size = new System.Drawing.Size(77, 13);
            this.lblFileForecastPath.TabIndex = 26;
            this.lblFileForecastPath.Text = "\"Forecast\" file:";
            // 
            // cmbFileSuperDettagliPath
            // 
            this.cmbFileSuperDettagliPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFileSuperDettagliPath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbFileSuperDettagliPath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbFileSuperDettagliPath.FormattingEnabled = true;
            this.cmbFileSuperDettagliPath.Location = new System.Drawing.Point(116, 96);
            this.cmbFileSuperDettagliPath.Name = "cmbFileSuperDettagliPath";
            this.cmbFileSuperDettagliPath.Size = new System.Drawing.Size(1357, 21);
            this.cmbFileSuperDettagliPath.TabIndex = 12;
            this.cmbFileSuperDettagliPath.SelectedIndexChanged += new System.EventHandler(this.cmbFileSuperDettagliPath_SelectedIndexChanged);
            this.cmbFileSuperDettagliPath.TextUpdate += new System.EventHandler(this.cmbFileSuperDettagliPath_TextUpdate);
            // 
            // lblFileSuperDettagliPath
            // 
            this.lblFileSuperDettagliPath.AutoSize = true;
            this.lblFileSuperDettagliPath.Location = new System.Drawing.Point(5, 98);
            this.lblFileSuperDettagliPath.Name = "lblFileSuperDettagliPath";
            this.lblFileSuperDettagliPath.Size = new System.Drawing.Size(101, 13);
            this.lblFileSuperDettagliPath.TabIndex = 31;
            this.lblFileSuperDettagliPath.Text = "\"Super dettagli\" file:";
            // 
            // cmbFileRanRatePath
            // 
            this.cmbFileRanRatePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFileRanRatePath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbFileRanRatePath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbFileRanRatePath.FormattingEnabled = true;
            this.cmbFileRanRatePath.Location = new System.Drawing.Point(115, 69);
            this.cmbFileRanRatePath.Name = "cmbFileRanRatePath";
            this.cmbFileRanRatePath.Size = new System.Drawing.Size(1357, 21);
            this.cmbFileRanRatePath.TabIndex = 8;
            this.cmbFileRanRatePath.SelectedIndexChanged += new System.EventHandler(this.cmbFileRanRatePath_SelectedIndexChanged);
            this.cmbFileRanRatePath.TextUpdate += new System.EventHandler(this.cmbFileRanRatePath_TextUpdate);
            // 
            // lblFileRanRatePath
            // 
            this.lblFileRanRatePath.AutoSize = true;
            this.lblFileRanRatePath.Location = new System.Drawing.Point(6, 71);
            this.lblFileRanRatePath.Name = "lblFileRanRatePath";
            this.lblFileRanRatePath.Size = new System.Drawing.Size(77, 13);
            this.lblFileRanRatePath.TabIndex = 36;
            this.lblFileRanRatePath.Text = "\"Ran rate\" file:";
            // 
            // btnCreaPresentazione
            // 
            this.btnCreaPresentazione.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCreaPresentazione.Location = new System.Drawing.Point(1464, 511);
            this.btnCreaPresentazione.Name = "btnCreaPresentazione";
            this.btnCreaPresentazione.Size = new System.Drawing.Size(120, 30);
            this.btnCreaPresentazione.TabIndex = 22;
            this.btnCreaPresentazione.Text = "Build &Presentation";
            this.btnCreaPresentazione.UseVisualStyleBackColor = true;
            this.btnCreaPresentazione.Click += new System.EventHandler(this.btnCreaPresentazione_Click);
            // 
            // lblFiltriApplicabili
            // 
            this.lblFiltriApplicabili.AutoSize = true;
            this.lblFiltriApplicabili.Location = new System.Drawing.Point(18, 57);
            this.lblFiltriApplicabili.Name = "lblFiltriApplicabili";
            this.lblFiltriApplicabili.Size = new System.Drawing.Size(37, 13);
            this.lblFiltriApplicabili.TabIndex = 39;
            this.lblFiltriApplicabili.Text = "Filters:";
            // 
            // dgvFiltri
            // 
            this.dgvFiltri.AllowUserToAddRows = false;
            this.dgvFiltri.AllowUserToDeleteRows = false;
            this.dgvFiltri.AllowUserToResizeRows = false;
            this.dgvFiltri.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvFiltri.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvFiltri.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgvFiltri.CausesValidation = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvFiltri.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvFiltri.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvFiltri.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Tabella,
            this.Campo,
            this.OpenFiltersSelection,
            this.ValoriSelezionati});
            this.dgvFiltri.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dgvFiltri.Location = new System.Drawing.Point(121, 47);
            this.dgvFiltri.Name = "dgvFiltri";
            this.dgvFiltri.RowHeadersVisible = false;
            this.dgvFiltri.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvFiltri.ShowEditingIcon = false;
            this.dgvFiltri.Size = new System.Drawing.Size(1454, 239);
            this.dgvFiltri.TabIndex = 21;
            // 
            // Tabella
            // 
            this.Tabella.HeaderText = "Table";
            this.Tabella.Name = "Tabella";
            this.Tabella.ReadOnly = true;
            this.Tabella.Width = 64;
            // 
            // Campo
            // 
            this.Campo.HeaderText = "Field";
            this.Campo.Name = "Campo";
            this.Campo.ReadOnly = true;
            this.Campo.Width = 59;
            // 
            // OpenFiltersSelection
            // 
            this.OpenFiltersSelection.HeaderText = "";
            this.OpenFiltersSelection.Name = "OpenFiltersSelection";
            this.OpenFiltersSelection.ReadOnly = true;
            this.OpenFiltersSelection.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.OpenFiltersSelection.Width = 5;
            // 
            // ValoriSelezionati
            // 
            this.ValoriSelezionati.HeaderText = "Selected values";
            this.ValoriSelezionati.Name = "ValoriSelezionati";
            this.ValoriSelezionati.ReadOnly = true;
            this.ValoriSelezionati.Width = 123;
            // 
            // btnClear
            // 
            this.btnClear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClear.AutoSize = true;
            this.btnClear.Location = new System.Drawing.Point(1458, 557);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(108, 13);
            this.btnClear.TabIndex = 24;
            this.btnClear.TabStop = true;
            this.btnClear.Text = "Clean current session";
            this.btnClear.Visible = false;
            this.btnClear.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.btnClear_LinkClicked);
            // 
            // calendarPeriodo
            // 
            this.calendarPeriodo.Location = new System.Drawing.Point(0, 0);
            this.calendarPeriodo.Name = "calendarPeriodo";
            this.calendarPeriodo.TabIndex = 20;
            this.calendarPeriodo.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.calendarPeriodo_DateSelected);
            // 
            // pnlCalendar
            // 
            this.pnlCalendar.Controls.Add(this.calendarPeriodo);
            this.pnlCalendar.Location = new System.Drawing.Point(267, 19);
            this.pnlCalendar.Name = "pnlCalendar";
            this.pnlCalendar.Size = new System.Drawing.Size(226, 161);
            this.pnlCalendar.TabIndex = 49;
            this.pnlCalendar.Visible = false;
            // 
            // gbPaths
            // 
            this.gbPaths.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbPaths.Controls.Add(this.lblFileBudgetPath);
            this.gbPaths.Controls.Add(this.btnSelectFileBudget);
            this.gbPaths.Controls.Add(this.cmbFileBudgetPath);
            this.gbPaths.Controls.Add(this.lblCartellaOutputPath);
            this.gbPaths.Controls.Add(this.btnSelectDestinationFolder);
            this.gbPaths.Controls.Add(this.cmbDestinationFolderPath);
            this.gbPaths.Controls.Add(this.btnOpenFileBudgetFolder);
            this.gbPaths.Controls.Add(this.btnOpenDestFolder);
            this.gbPaths.Controls.Add(this.btnOpenFileRanRate);
            this.gbPaths.Controls.Add(this.btnOpenFileBudget);
            this.gbPaths.Controls.Add(this.btnOpenFileRanRateFolder);
            this.gbPaths.Controls.Add(this.lblFileForecastPath);
            this.gbPaths.Controls.Add(this.cmbFileRanRatePath);
            this.gbPaths.Controls.Add(this.btnSelectForecastFile);
            this.gbPaths.Controls.Add(this.btnSelectFileRanRate);
            this.gbPaths.Controls.Add(this.cmbFileForecastPath);
            this.gbPaths.Controls.Add(this.lblFileRanRatePath);
            this.gbPaths.Controls.Add(this.btnOpenFileForecastFolder);
            this.gbPaths.Controls.Add(this.btnOpenFileSuperDettagli);
            this.gbPaths.Controls.Add(this.btnOpenFileForecast);
            this.gbPaths.Controls.Add(this.btnOpenFileSuperDettagliFolder);
            this.gbPaths.Controls.Add(this.lblFileSuperDettagliPath);
            this.gbPaths.Controls.Add(this.cmbFileSuperDettagliPath);
            this.gbPaths.Controls.Add(this.btnSelectFileSuperDettagli);
            this.gbPaths.Location = new System.Drawing.Point(9, 27);
            this.gbPaths.Name = "gbPaths";
            this.gbPaths.Size = new System.Drawing.Size(1581, 151);
            this.gbPaths.TabIndex = 50;
            this.gbPaths.TabStop = false;
            this.gbPaths.Text = "Paths";
            // 
            // gbOptions
            // 
            this.gbOptions.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbOptions.Controls.Add(this.lblDataPeriodo);
            this.gbOptions.Controls.Add(this.lblFiltriApplicabili);
            this.gbOptions.Controls.Add(this.pnlCalendar);
            this.gbOptions.Controls.Add(this.dgvFiltri);
            this.gbOptions.Controls.Add(this.label4);
            this.gbOptions.Controls.Add(this.btnOpenCalendar);
            this.gbOptions.Location = new System.Drawing.Point(9, 213);
            this.gbOptions.Name = "gbOptions";
            this.gbOptions.Size = new System.Drawing.Size(1581, 292);
            this.gbOptions.TabIndex = 51;
            this.gbOptions.TabStop = false;
            this.gbOptions.Text = "Options";
            // 
            // btnNext
            // 
            this.btnNext.Location = new System.Drawing.Point(1470, 184);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(120, 30);
            this.btnNext.TabIndex = 52;
            this.btnNext.Text = "&Next >>";
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // bfbDestFolder
            // 
            this.bfbDestFolder.Multiselect = false;
            this.bfbDestFolder.RootFolder = "C:\\Users\\miche\\Desktop";
            this.bfbDestFolder.Title = "Please select a folder...";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1602, 1035);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.gbOptions);
            this.Controls.Add(this.gbPaths);
            this.Controls.Add(this.btnCreaPresentazione);
            this.Controls.Add(this.btnCopyError);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.wbExecutionResult);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.statusStrip);
            this.Controls.Add(this.menuStrip);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip;
            this.MinimumSize = new System.Drawing.Size(1000, 650);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PowerPoint Generator";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvFiltri)).EndInit();
            this.pnlCalendar.ResumeLayout(false);
            this.gbPaths.ResumeLayout(false);
            this.gbPaths.PerformLayout();
            this.gbOptions.ResumeLayout(false);
            this.gbOptions.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel txtStatusLabel;
        private System.Windows.Forms.ToolTip toolTipDefault;
        private System.Windows.Forms.ComboBox cmbFileBudgetPath;
        private System.Windows.Forms.Button btnSelectFileBudget;
        private System.Windows.Forms.Label lblFileBudgetPath;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.MenuStrip menuStrip;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemClear;
        private System.Windows.Forms.ComboBox cmbDestinationFolderPath;
        private System.Windows.Forms.Button btnSelectDestinationFolder;
        private System.Windows.Forms.Label lblCartellaOutputPath;
        private System.Windows.Forms.Button btnOpenFileBudgetFolder;
        private System.Windows.Forms.Button btnOpenDestFolder;
        private System.Windows.Forms.Button btnOpenFileBudget;
        private System.Windows.Forms.WebBrowser wbExecutionResult;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.ComponentModel.BackgroundWorker btnCreatePresentationBackgroundWorker;
        private System.Windows.Forms.LinkLabel btnCopyError;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar;
        private WK.Libraries.BetterFolderBrowserNS.BetterFolderBrowser bfbDestFolder;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        private System.Windows.Forms.ToolStripStatusLabel lblVersion;
        private System.Windows.Forms.Button btnOpenFileForecast;
        private System.Windows.Forms.Button btnOpenFileForecastFolder;
        private System.Windows.Forms.ComboBox cmbFileForecastPath;
        private System.Windows.Forms.Button btnSelectForecastFile;
        private System.Windows.Forms.Label lblFileForecastPath;
        private System.Windows.Forms.Button btnOpenFileSuperDettagli;
        private System.Windows.Forms.Button btnOpenFileSuperDettagliFolder;
        private System.Windows.Forms.ComboBox cmbFileSuperDettagliPath;
        private System.Windows.Forms.Button btnSelectFileSuperDettagli;
        private System.Windows.Forms.Label lblFileSuperDettagliPath;
        private System.Windows.Forms.Button btnOpenFileRanRate;
        private System.Windows.Forms.Button btnOpenFileRanRateFolder;
        private System.Windows.Forms.ComboBox cmbFileRanRatePath;
        private System.Windows.Forms.Button btnSelectFileRanRate;
        private System.Windows.Forms.Label lblFileRanRatePath;
        private System.Windows.Forms.Button btnCreaPresentazione;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemOpenConfigFolder;
        private System.Windows.Forms.Label lblFiltriApplicabili;
        private System.Windows.Forms.DataGridView dgvFiltri;
        private System.Windows.Forms.LinkLabel btnClear;
        private System.Windows.Forms.Label lblDataPeriodo;
        private System.Windows.Forms.Button btnOpenCalendar;
        private System.Windows.Forms.MonthCalendar calendarPeriodo;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel pnlCalendar;
        private System.Windows.Forms.GroupBox gbPaths;
        private System.Windows.Forms.GroupBox gbOptions;
        private System.Windows.Forms.Button btnNext;
        private System.ComponentModel.BackgroundWorker btnNextBackgroundWorker;
        private System.Windows.Forms.DataGridViewTextBoxColumn Tabella;
        private System.Windows.Forms.DataGridViewTextBoxColumn Campo;
        private System.Windows.Forms.DataGridViewButtonColumn OpenFiltersSelection;
        private System.Windows.Forms.DataGridViewTextBoxColumn ValoriSelezionati;
    }
}

