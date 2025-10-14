namespace PptGeneratorGUI
{
    partial class frmMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
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
            this.btnSelectFileRunRate = new System.Windows.Forms.Button();
            this.lblDataPeriodo = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnOpenCalendar = new System.Windows.Forms.Button();
            this.btnOpenFileRunRate = new System.Windows.Forms.Button();
            this.btnOpenFileRunRateFolder = new System.Windows.Forms.Button();
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
            this.cmbDestinationFolderPath = new System.Windows.Forms.ComboBox();
            this.wbExecutionResult = new System.Windows.Forms.WebBrowser();
            this.buildPresentationBackgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.btnCopyError = new System.Windows.Forms.LinkLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.cmbFileForecastPath = new System.Windows.Forms.ComboBox();
            this.lblFileForecastPath = new System.Windows.Forms.Label();
            this.cmbFileSuperDettagliPath = new System.Windows.Forms.ComboBox();
            this.lblFileSuperDettagliPath = new System.Windows.Forms.Label();
            this.cmbFileRunRatePath = new System.Windows.Forms.ComboBox();
            this.lblFileRunRatePath = new System.Windows.Forms.Label();
            this.btnBuildPresentation = new System.Windows.Forms.Button();
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
            this.cbReplaceAllDataFileSuperDettagli = new System.Windows.Forms.CheckBox();
            this.gbOptions = new System.Windows.Forms.GroupBox();
            this.btnValidaInput = new System.Windows.Forms.Button();
            this.validaInputBackgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.clearPathsHistoryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openSouceFilesFolderToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.bfbDestFolder = new WK.Libraries.BetterFolderBrowserNS.BetterFolderBrowser(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.dgvFiltri)).BeginInit();
            this.pnlCalendar.SuspendLayout();
            this.gbPaths.SuspendLayout();
            this.gbOptions.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // statusStrip
            // 
            this.statusStrip.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.statusStrip.Location = new System.Drawing.Point(0, 902);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Padding = new System.Windows.Forms.Padding(2, 0, 14, 0);
            this.statusStrip.Size = new System.Drawing.Size(1178, 22);
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
            this.btnSelectFileBudget.Location = new System.Drawing.Point(1057, 14);
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
            this.btnSelectDestinationFolder.Location = new System.Drawing.Point(1057, 120);
            this.btnSelectDestinationFolder.Name = "btnSelectDestinationFolder";
            this.btnSelectDestinationFolder.Size = new System.Drawing.Size(29, 23);
            this.btnSelectDestinationFolder.TabIndex = 18;
            this.btnSelectDestinationFolder.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectDestinationFolder, "Apre la finestra di selezione cartella");
            this.btnSelectDestinationFolder.UseVisualStyleBackColor = true;
            this.btnSelectDestinationFolder.Click += new System.EventHandler(this.btnSelectDestinationFolder_Click);
            // 
            // lblCartellaOutputPath
            // 
            this.lblCartellaOutputPath.AutoSize = true;
            this.lblCartellaOutputPath.Location = new System.Drawing.Point(3, 123);
            this.lblCartellaOutputPath.Name = "lblCartellaOutputPath";
            this.lblCartellaOutputPath.Size = new System.Drawing.Size(71, 13);
            this.lblCartellaOutputPath.TabIndex = 11;
            this.lblCartellaOutputPath.Text = "Output folder:";
            this.toolTipDefault.SetToolTip(this.lblCartellaOutputPath, "Cartella nella quale vettanno salvati i file di output");
            // 
            // btnSelectForecastFile
            // 
            this.btnSelectForecastFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectForecastFile.Location = new System.Drawing.Point(1057, 40);
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
            this.btnSelectFileSuperDettagli.Location = new System.Drawing.Point(1057, 94);
            this.btnSelectFileSuperDettagli.Name = "btnSelectFileSuperDettagli";
            this.btnSelectFileSuperDettagli.Size = new System.Drawing.Size(29, 23);
            this.btnSelectFileSuperDettagli.TabIndex = 14;
            this.btnSelectFileSuperDettagli.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectFileSuperDettagli, "Apre la finestra di selezione file");
            this.btnSelectFileSuperDettagli.UseVisualStyleBackColor = true;
            this.btnSelectFileSuperDettagli.Click += new System.EventHandler(this.btnSelectFileSuperDettagli_Click);
            // 
            // btnSelectFileRunRate
            // 
            this.btnSelectFileRunRate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFileRunRate.Location = new System.Drawing.Point(1057, 67);
            this.btnSelectFileRunRate.Name = "btnSelectFileRunRate";
            this.btnSelectFileRunRate.Size = new System.Drawing.Size(29, 23);
            this.btnSelectFileRunRate.TabIndex = 9;
            this.btnSelectFileRunRate.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectFileRunRate, "Apre la finestra di selezione file");
            this.btnSelectFileRunRate.UseVisualStyleBackColor = true;
            this.btnSelectFileRunRate.Click += new System.EventHandler(this.btnSelectFileRunRate_Click);
            // 
            // lblDataPeriodo
            // 
            this.lblDataPeriodo.AutoSize = true;
            this.lblDataPeriodo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDataPeriodo.Location = new System.Drawing.Point(102, 26);
            this.lblDataPeriodo.Name = "lblDataPeriodo";
            this.lblDataPeriodo.Size = new System.Drawing.Size(108, 13);
            this.lblDataPeriodo.TabIndex = 21;
            this.lblDataPeriodo.Text = "[DATA PERIODO]";
            this.toolTipDefault.SetToolTip(this.lblDataPeriodo, "Cartella nella quale vettanno salvati i file di output");
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 23);
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
            this.btnOpenCalendar.TabIndex = 22;
            this.toolTipDefault.SetToolTip(this.btnOpenCalendar, "Apri il calendario per selezionare la data del periodo di riferimento");
            this.btnOpenCalendar.UseVisualStyleBackColor = true;
            this.btnOpenCalendar.Click += new System.EventHandler(this.btnOpenCalendar_Click);
            // 
            // btnOpenFileRunRate
            // 
            this.btnOpenFileRunRate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileRunRate.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileRunRate.Image")));
            this.btnOpenFileRunRate.Location = new System.Drawing.Point(1128, 66);
            this.btnOpenFileRunRate.Name = "btnOpenFileRunRate";
            this.btnOpenFileRunRate.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileRunRate.TabIndex = 11;
            this.toolTipDefault.SetToolTip(this.btnOpenFileRunRate, "Apre il file report");
            this.btnOpenFileRunRate.UseVisualStyleBackColor = true;
            this.btnOpenFileRunRate.Click += new System.EventHandler(this.btnOpenFileRunRate_Click);
            // 
            // btnOpenFileRunRateFolder
            // 
            this.btnOpenFileRunRateFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileRunRateFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileRunRateFolder.Image")));
            this.btnOpenFileRunRateFolder.Location = new System.Drawing.Point(1092, 66);
            this.btnOpenFileRunRateFolder.Name = "btnOpenFileRunRateFolder";
            this.btnOpenFileRunRateFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileRunRateFolder.TabIndex = 10;
            this.toolTipDefault.SetToolTip(this.btnOpenFileRunRateFolder, "Apre la cartella contenente il file report");
            this.btnOpenFileRunRateFolder.UseVisualStyleBackColor = true;
            this.btnOpenFileRunRateFolder.Click += new System.EventHandler(this.btnOpenFileRunRateFolder_Click);
            // 
            // btnOpenFileSuperDettagli
            // 
            this.btnOpenFileSuperDettagli.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileSuperDettagli.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileSuperDettagli.Image")));
            this.btnOpenFileSuperDettagli.Location = new System.Drawing.Point(1127, 93);
            this.btnOpenFileSuperDettagli.Name = "btnOpenFileSuperDettagli";
            this.btnOpenFileSuperDettagli.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileSuperDettagli.TabIndex = 16;
            this.toolTipDefault.SetToolTip(this.btnOpenFileSuperDettagli, "Apre il file report");
            this.btnOpenFileSuperDettagli.UseVisualStyleBackColor = true;
            this.btnOpenFileSuperDettagli.Click += new System.EventHandler(this.btnOpenFileSuperDettagli_Click);
            // 
            // btnOpenFileSuperDettagliFolder
            // 
            this.btnOpenFileSuperDettagliFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileSuperDettagliFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileSuperDettagliFolder.Image")));
            this.btnOpenFileSuperDettagliFolder.Location = new System.Drawing.Point(1092, 94);
            this.btnOpenFileSuperDettagliFolder.Name = "btnOpenFileSuperDettagliFolder";
            this.btnOpenFileSuperDettagliFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileSuperDettagliFolder.TabIndex = 15;
            this.toolTipDefault.SetToolTip(this.btnOpenFileSuperDettagliFolder, "Apre la cartella contenente il file report");
            this.btnOpenFileSuperDettagliFolder.UseVisualStyleBackColor = true;
            this.btnOpenFileSuperDettagliFolder.Click += new System.EventHandler(this.btnOpenFileSuperDettagliFolder_Click);
            // 
            // btnOpenFileForecast
            // 
            this.btnOpenFileForecast.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileForecast.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileForecast.Image")));
            this.btnOpenFileForecast.Location = new System.Drawing.Point(1128, 40);
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
            this.btnOpenFileForecastFolder.Location = new System.Drawing.Point(1092, 40);
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
            this.btnOpenFileBudget.Location = new System.Drawing.Point(1128, 14);
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
            this.btnOpenDestFolder.Location = new System.Drawing.Point(1092, 121);
            this.btnOpenDestFolder.Name = "btnOpenDestFolder";
            this.btnOpenDestFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenDestFolder.TabIndex = 19;
            this.toolTipDefault.SetToolTip(this.btnOpenDestFolder, "Apre la cartella di destinazione");
            this.btnOpenDestFolder.UseVisualStyleBackColor = true;
            this.btnOpenDestFolder.Click += new System.EventHandler(this.btnOpenDestFolder_Click);
            // 
            // btnOpenFileBudgetFolder
            // 
            this.btnOpenFileBudgetFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileBudgetFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileBudgetFolder.Image")));
            this.btnOpenFileBudgetFolder.Location = new System.Drawing.Point(1092, 14);
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
            this.cmbFileBudgetPath.Location = new System.Drawing.Point(104, 16);
            this.cmbFileBudgetPath.Name = "cmbFileBudgetPath";
            this.cmbFileBudgetPath.Size = new System.Drawing.Size(947, 21);
            this.cmbFileBudgetPath.TabIndex = 0;
            this.cmbFileBudgetPath.SelectedIndexChanged += new System.EventHandler(this.cmbFileBudgetPath_SelectedIndexChanged);
            this.cmbFileBudgetPath.TextUpdate += new System.EventHandler(this.cmbFileBudgetPath_TextUpdate);
            // 
            // lblFileBudgetPath
            // 
            this.lblFileBudgetPath.AutoSize = true;
            this.lblFileBudgetPath.Location = new System.Drawing.Point(3, 19);
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
            // cmbDestinationFolderPath
            // 
            this.cmbDestinationFolderPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbDestinationFolderPath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbDestinationFolderPath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbDestinationFolderPath.FormattingEnabled = true;
            this.cmbDestinationFolderPath.Location = new System.Drawing.Point(104, 120);
            this.cmbDestinationFolderPath.Name = "cmbDestinationFolderPath";
            this.cmbDestinationFolderPath.Size = new System.Drawing.Size(947, 21);
            this.cmbDestinationFolderPath.TabIndex = 17;
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
            this.wbExecutionResult.Size = new System.Drawing.Size(1154, 352);
            this.wbExecutionResult.TabIndex = 25;
            this.wbExecutionResult.Navigating += new System.Windows.Forms.WebBrowserNavigatingEventHandler(this.wbExecutionResult_Navigating_1);
            // 
            // buildPresentationBackgroundWorker
            // 
            this.buildPresentationBackgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.buildPresentationBackgroundWorker_DoWork);
            this.buildPresentationBackgroundWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.buildPresentationBackgroundWorker_RunWorkerCompleted);
            // 
            // btnCopyError
            // 
            this.btnCopyError.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCopyError.AutoSize = true;
            this.btnCopyError.Location = new System.Drawing.Point(994, 988);
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
            this.cmbFileForecastPath.Location = new System.Drawing.Point(104, 42);
            this.cmbFileForecastPath.Name = "cmbFileForecastPath";
            this.cmbFileForecastPath.Size = new System.Drawing.Size(947, 21);
            this.cmbFileForecastPath.TabIndex = 4;
            this.cmbFileForecastPath.SelectedIndexChanged += new System.EventHandler(this.cmbFileForecastPath_SelectedIndexChanged);
            this.cmbFileForecastPath.TextUpdate += new System.EventHandler(this.cmbFileForecastPath_TextUpdate);
            // 
            // lblFileForecastPath
            // 
            this.lblFileForecastPath.AutoSize = true;
            this.lblFileForecastPath.Location = new System.Drawing.Point(3, 45);
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
            this.cmbFileSuperDettagliPath.Location = new System.Drawing.Point(104, 94);
            this.cmbFileSuperDettagliPath.Name = "cmbFileSuperDettagliPath";
            this.cmbFileSuperDettagliPath.Size = new System.Drawing.Size(848, 21);
            this.cmbFileSuperDettagliPath.TabIndex = 12;
            this.cmbFileSuperDettagliPath.SelectedIndexChanged += new System.EventHandler(this.cmbFileSuperDettagliPath_SelectedIndexChanged);
            this.cmbFileSuperDettagliPath.TextUpdate += new System.EventHandler(this.cmbFileSuperDettagliPath_TextUpdate);
            // 
            // lblFileSuperDettagliPath
            // 
            this.lblFileSuperDettagliPath.AutoSize = true;
            this.lblFileSuperDettagliPath.Location = new System.Drawing.Point(3, 97);
            this.lblFileSuperDettagliPath.Name = "lblFileSuperDettagliPath";
            this.lblFileSuperDettagliPath.Size = new System.Drawing.Size(101, 13);
            this.lblFileSuperDettagliPath.TabIndex = 31;
            this.lblFileSuperDettagliPath.Text = "\"Super dettagli\" file:";
            // 
            // cmbFileRunRatePath
            // 
            this.cmbFileRunRatePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFileRunRatePath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbFileRunRatePath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbFileRunRatePath.FormattingEnabled = true;
            this.cmbFileRunRatePath.Location = new System.Drawing.Point(104, 68);
            this.cmbFileRunRatePath.Name = "cmbFileRunRatePath";
            this.cmbFileRunRatePath.Size = new System.Drawing.Size(947, 21);
            this.cmbFileRunRatePath.TabIndex = 8;
            this.cmbFileRunRatePath.SelectedIndexChanged += new System.EventHandler(this.cmbFileRunRatePath_SelectedIndexChanged);
            this.cmbFileRunRatePath.TextUpdate += new System.EventHandler(this.cmbFileRunRatePath_TextUpdate);
            // 
            // lblFileRunRatePath
            // 
            this.lblFileRunRatePath.AutoSize = true;
            this.lblFileRunRatePath.Location = new System.Drawing.Point(3, 71);
            this.lblFileRunRatePath.Name = "lblFileRunRatePath";
            this.lblFileRunRatePath.Size = new System.Drawing.Size(77, 13);
            this.lblFileRunRatePath.TabIndex = 36;
            this.lblFileRunRatePath.Text = "\"Run rate\" file:";
            // 
            // btnBuildPresentation
            // 
            this.btnBuildPresentation.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBuildPresentation.Location = new System.Drawing.Point(1040, 511);
            this.btnBuildPresentation.Name = "btnBuildPresentation";
            this.btnBuildPresentation.Size = new System.Drawing.Size(120, 30);
            this.btnBuildPresentation.TabIndex = 24;
            this.btnBuildPresentation.Text = "Build &Presentation";
            this.btnBuildPresentation.UseVisualStyleBackColor = true;
            this.btnBuildPresentation.Click += new System.EventHandler(this.btnBuildPresentation_Click);
            // 
            // lblFiltriApplicabili
            // 
            this.lblFiltriApplicabili.AutoSize = true;
            this.lblFiltriApplicabili.Location = new System.Drawing.Point(3, 57);
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
            this.dgvFiltri.Location = new System.Drawing.Point(104, 47);
            this.dgvFiltri.Name = "dgvFiltri";
            this.dgvFiltri.RowHeadersVisible = false;
            this.dgvFiltri.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvFiltri.ShowEditingIcon = false;
            this.dgvFiltri.Size = new System.Drawing.Size(1047, 239);
            this.dgvFiltri.TabIndex = 23;
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
            this.btnClear.Location = new System.Drawing.Point(1034, 557);
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
            this.gbPaths.Controls.Add(this.cbReplaceAllDataFileSuperDettagli);
            this.gbPaths.Controls.Add(this.lblFileBudgetPath);
            this.gbPaths.Controls.Add(this.btnSelectFileBudget);
            this.gbPaths.Controls.Add(this.cmbFileBudgetPath);
            this.gbPaths.Controls.Add(this.lblCartellaOutputPath);
            this.gbPaths.Controls.Add(this.btnSelectDestinationFolder);
            this.gbPaths.Controls.Add(this.cmbDestinationFolderPath);
            this.gbPaths.Controls.Add(this.btnOpenFileBudgetFolder);
            this.gbPaths.Controls.Add(this.btnOpenDestFolder);
            this.gbPaths.Controls.Add(this.btnOpenFileRunRate);
            this.gbPaths.Controls.Add(this.btnOpenFileBudget);
            this.gbPaths.Controls.Add(this.btnOpenFileRunRateFolder);
            this.gbPaths.Controls.Add(this.lblFileForecastPath);
            this.gbPaths.Controls.Add(this.cmbFileRunRatePath);
            this.gbPaths.Controls.Add(this.btnSelectForecastFile);
            this.gbPaths.Controls.Add(this.btnSelectFileRunRate);
            this.gbPaths.Controls.Add(this.cmbFileForecastPath);
            this.gbPaths.Controls.Add(this.lblFileRunRatePath);
            this.gbPaths.Controls.Add(this.btnOpenFileForecastFolder);
            this.gbPaths.Controls.Add(this.btnOpenFileSuperDettagli);
            this.gbPaths.Controls.Add(this.btnOpenFileForecast);
            this.gbPaths.Controls.Add(this.btnOpenFileSuperDettagliFolder);
            this.gbPaths.Controls.Add(this.lblFileSuperDettagliPath);
            this.gbPaths.Controls.Add(this.cmbFileSuperDettagliPath);
            this.gbPaths.Controls.Add(this.btnSelectFileSuperDettagli);
            this.gbPaths.Location = new System.Drawing.Point(3, 27);
            this.gbPaths.Name = "gbPaths";
            this.gbPaths.Size = new System.Drawing.Size(1157, 151);
            this.gbPaths.TabIndex = 50;
            this.gbPaths.TabStop = false;
            this.gbPaths.Text = "Paths";
            // 
            // cbReplaceAllDataFileSuperDettagli
            // 
            this.cbReplaceAllDataFileSuperDettagli.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cbReplaceAllDataFileSuperDettagli.AutoSize = true;
            this.cbReplaceAllDataFileSuperDettagli.Location = new System.Drawing.Point(961, 97);
            this.cbReplaceAllDataFileSuperDettagli.Name = "cbReplaceAllDataFileSuperDettagli";
            this.cbReplaceAllDataFileSuperDettagli.Size = new System.Drawing.Size(90, 17);
            this.cbReplaceAllDataFileSuperDettagli.TabIndex = 13;
            this.cbReplaceAllDataFileSuperDettagli.Text = "Replace data";
            this.cbReplaceAllDataFileSuperDettagli.UseVisualStyleBackColor = true;
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
            this.gbOptions.Location = new System.Drawing.Point(3, 213);
            this.gbOptions.Name = "gbOptions";
            this.gbOptions.Size = new System.Drawing.Size(1157, 292);
            this.gbOptions.TabIndex = 51;
            this.gbOptions.TabStop = false;
            this.gbOptions.Text = "Options";
            // 
            // btnValidaInput
            // 
            this.btnValidaInput.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnValidaInput.Location = new System.Drawing.Point(1040, 184);
            this.btnValidaInput.Name = "btnValidaInput";
            this.btnValidaInput.Size = new System.Drawing.Size(120, 30);
            this.btnValidaInput.TabIndex = 20;
            this.btnValidaInput.Text = "&Validate input";
            this.btnValidaInput.UseVisualStyleBackColor = true;
            this.btnValidaInput.Click += new System.EventHandler(this.btnValidaInput_Click);
            // 
            // validaInputBackgroundWorker
            // 
            this.validaInputBackgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.validaInputBackgroundWorker_DoWork);
            this.validaInputBackgroundWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.validaInputBackgroundWorker_RunWorkerCompleted);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1178, 24);
            this.menuStrip1.TabIndex = 53;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.clearPathsHistoryToolStripMenuItem,
            this.openSouceFilesFolderToolStripMenuItem});
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(61, 20);
            this.toolStripMenuItem1.Text = "&Options";
            // 
            // clearPathsHistoryToolStripMenuItem
            // 
            this.clearPathsHistoryToolStripMenuItem.Name = "clearPathsHistoryToolStripMenuItem";
            this.clearPathsHistoryToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.clearPathsHistoryToolStripMenuItem.Text = "&Clear paths history";
            this.clearPathsHistoryToolStripMenuItem.Click += new System.EventHandler(this.clearPathsHistoryToolStripMenuItem_Click);
            // 
            // openSouceFilesFolderToolStripMenuItem
            // 
            this.openSouceFilesFolderToolStripMenuItem.Name = "openSouceFilesFolderToolStripMenuItem";
            this.openSouceFilesFolderToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.openSouceFilesFolderToolStripMenuItem.Text = "Open &Souce files folder";
            this.openSouceFilesFolderToolStripMenuItem.Click += new System.EventHandler(this.openSouceFilesFolderToolStripMenuItem_Click);
            // 
            // bfbDestFolder
            // 
            this.bfbDestFolder.Multiselect = false;
            this.bfbDestFolder.RootFolder = "C:\\Users\\miche\\Desktop";
            this.bfbDestFolder.Title = "Please select a folder...";
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1178, 924);
            this.Controls.Add(this.btnValidaInput);
            this.Controls.Add(this.gbOptions);
            this.Controls.Add(this.gbPaths);
            this.Controls.Add(this.btnBuildPresentation);
            this.Controls.Add(this.btnCopyError);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.wbExecutionResult);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.statusStrip);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(1000, 650);
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PowerPoint Generator";
            this.Load += new System.EventHandler(this.MainForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvFiltri)).EndInit();
            this.pnlCalendar.ResumeLayout(false);
            this.gbPaths.ResumeLayout(false);
            this.gbPaths.PerformLayout();
            this.gbOptions.ResumeLayout(false);
            this.gbOptions.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
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
        private System.Windows.Forms.ComboBox cmbDestinationFolderPath;
        private System.Windows.Forms.Button btnSelectDestinationFolder;
        private System.Windows.Forms.Label lblCartellaOutputPath;
        private System.Windows.Forms.Button btnOpenFileBudgetFolder;
        private System.Windows.Forms.Button btnOpenDestFolder;
        private System.Windows.Forms.Button btnOpenFileBudget;
        private System.Windows.Forms.WebBrowser wbExecutionResult;
        private System.ComponentModel.BackgroundWorker buildPresentationBackgroundWorker;
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
        private System.Windows.Forms.Button btnOpenFileRunRate;
        private System.Windows.Forms.Button btnOpenFileRunRateFolder;
        private System.Windows.Forms.ComboBox cmbFileRunRatePath;
        private System.Windows.Forms.Button btnSelectFileRunRate;
        private System.Windows.Forms.Label lblFileRunRatePath;
        private System.Windows.Forms.Button btnBuildPresentation;
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
        private System.Windows.Forms.Button btnValidaInput;
        private System.ComponentModel.BackgroundWorker validaInputBackgroundWorker;
        private System.Windows.Forms.DataGridViewTextBoxColumn Tabella;
        private System.Windows.Forms.DataGridViewTextBoxColumn Campo;
        private System.Windows.Forms.DataGridViewButtonColumn OpenFiltersSelection;
        private System.Windows.Forms.DataGridViewTextBoxColumn ValoriSelezionati;
        private System.Windows.Forms.CheckBox cbReplaceAllDataFileSuperDettagli;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem clearPathsHistoryToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openSouceFilesFolderToolStripMenuItem;
    }
}

