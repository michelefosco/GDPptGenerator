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
            this.btnOpenFileRunRate = new System.Windows.Forms.Button();
            this.btnOpenFileRunRateFolder = new System.Windows.Forms.Button();
            this.btnOpenFileSuperDettagli = new System.Windows.Forms.Button();
            this.btnOpenFileSuperDettagliFolder = new System.Windows.Forms.Button();
            this.btnOpenFileForecast = new System.Windows.Forms.Button();
            this.btnOpenFileForecastFolder = new System.Windows.Forms.Button();
            this.btnOpenFileBudget = new System.Windows.Forms.Button();
            this.btnOpenDestFolder = new System.Windows.Forms.Button();
            this.btnOpenFileBudgetFolder = new System.Windows.Forms.Button();
            this.btnBuildPresentation = new System.Windows.Forms.Button();
            this.btnValidaInput = new System.Windows.Forms.Button();
            this.btnOpenFileCN43N = new System.Windows.Forms.Button();
            this.btnOpenFileCN43NFolder = new System.Windows.Forms.Button();
            this.btnSelectFileCN43N = new System.Windows.Forms.Button();
            this.lblDataPeriodo = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnOpenCalendar = new System.Windows.Forms.Button();
            this.cmbFileBudgetPath = new System.Windows.Forms.ComboBox();
            this.lblFileBudgetPath = new System.Windows.Forms.Label();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.lblResults = new System.Windows.Forms.Label();
            this.cmbDestinationFolderPath = new System.Windows.Forms.ComboBox();
            this.wbExecutionResult = new System.Windows.Forms.WebBrowser();
            this.btnCopyOutput = new System.Windows.Forms.LinkLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.cmbFileForecastPath = new System.Windows.Forms.ComboBox();
            this.lblFileForecastPath = new System.Windows.Forms.Label();
            this.cmbFileSuperDettagliPath = new System.Windows.Forms.ComboBox();
            this.lblFileSuperDettagliPath = new System.Windows.Forms.Label();
            this.cmbFileRunRatePath = new System.Windows.Forms.ComboBox();
            this.lblFileRunRatePath = new System.Windows.Forms.Label();
            this.lblFiltriApplicabili = new System.Windows.Forms.Label();
            this.dgvFiltri = new System.Windows.Forms.DataGridView();
            this.Table = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OpenFiltersSelection = new System.Windows.Forms.DataGridViewButtonColumn();
            this.SelectedValues = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.calendarPeriodo = new System.Windows.Forms.MonthCalendar();
            this.pnlCalendar = new System.Windows.Forms.Panel();
            this.gbPaths = new System.Windows.Forms.GroupBox();
            this.gbCN43N = new System.Windows.Forms.GroupBox();
            this.rbCN43N_OverwriteAll = new System.Windows.Forms.RadioButton();
            this.rbCN43N_Append = new System.Windows.Forms.RadioButton();
            this.btnGetWbsList = new System.Windows.Forms.Button();
            this.gbSuperDettagli = new System.Windows.Forms.GroupBox();
            this.rbSuperDettagli_Add = new System.Windows.Forms.RadioButton();
            this.rbSuperDettagli_ReplaceCurrentYear = new System.Windows.Forms.RadioButton();
            this.cmbFileCN43NPath = new System.Windows.Forms.ComboBox();
            this.lblFileCN43NPath = new System.Windows.Forms.Label();
            this.gbOptions = new System.Windows.Forms.GroupBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.openDataSouceExcelFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openSouceFilesFolderToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.sessionToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.loadLastSessionPathsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cleanCurrentsessionToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.deletePathsHistoryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.updatePresentationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.updatePresentationToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.lblElaborazioneInCorso = new System.Windows.Forms.Label();
            this.btnTryBuildPresentationOnly = new System.Windows.Forms.Button();
            this.bfbDestFolder = new WK.Libraries.BetterFolderBrowserNS.BetterFolderBrowser(this.components);
            this.statusStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvFiltri)).BeginInit();
            this.pnlCalendar.SuspendLayout();
            this.gbPaths.SuspendLayout();
            this.gbCN43N.SuspendLayout();
            this.gbSuperDettagli.SuspendLayout();
            this.gbOptions.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // statusStrip
            // 
            this.statusStrip.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripProgressBar,
            this.toolStripStatusLabel1,
            this.txtStatusLabel,
            this.lblVersion});
            this.statusStrip.Location = new System.Drawing.Point(0, 902);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Padding = new System.Windows.Forms.Padding(2, 0, 14, 0);
            this.statusStrip.Size = new System.Drawing.Size(1185, 22);
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
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(48, 17);
            this.toolStripStatusLabel1.Text = "  Status:";
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
            this.lblVersion.Size = new System.Drawing.Size(1072, 17);
            this.lblVersion.Spring = true;
            this.lblVersion.Text = "[VERSIONE]";
            this.lblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnSelectFileBudget
            // 
            this.btnSelectFileBudget.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFileBudget.Location = new System.Drawing.Point(1071, 14);
            this.btnSelectFileBudget.Name = "btnSelectFileBudget";
            this.btnSelectFileBudget.Size = new System.Drawing.Size(29, 23);
            this.btnSelectFileBudget.TabIndex = 1;
            this.btnSelectFileBudget.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectFileBudget, "Open the window to select a file");
            this.btnSelectFileBudget.UseVisualStyleBackColor = true;
            this.btnSelectFileBudget.Click += new System.EventHandler(this.btnSelectFileBudget_Click);
            // 
            // btnSelectDestinationFolder
            // 
            this.btnSelectDestinationFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectDestinationFolder.Location = new System.Drawing.Point(1071, 148);
            this.btnSelectDestinationFolder.Name = "btnSelectDestinationFolder";
            this.btnSelectDestinationFolder.Size = new System.Drawing.Size(29, 23);
            this.btnSelectDestinationFolder.TabIndex = 26;
            this.btnSelectDestinationFolder.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectDestinationFolder, "Open the window to select a file");
            this.btnSelectDestinationFolder.UseVisualStyleBackColor = true;
            this.btnSelectDestinationFolder.Click += new System.EventHandler(this.btnSelectDestinationFolder_Click);
            // 
            // lblCartellaOutputPath
            // 
            this.lblCartellaOutputPath.AutoSize = true;
            this.lblCartellaOutputPath.Location = new System.Drawing.Point(4, 151);
            this.lblCartellaOutputPath.Name = "lblCartellaOutputPath";
            this.lblCartellaOutputPath.Size = new System.Drawing.Size(75, 13);
            this.lblCartellaOutputPath.TabIndex = 11;
            this.lblCartellaOutputPath.Text = "Output folder*:";
            this.toolTipDefault.SetToolTip(this.lblCartellaOutputPath, "Folder where the output files will be saved into");
            // 
            // btnSelectForecastFile
            // 
            this.btnSelectForecastFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectForecastFile.Location = new System.Drawing.Point(1071, 40);
            this.btnSelectForecastFile.Name = "btnSelectForecastFile";
            this.btnSelectForecastFile.Size = new System.Drawing.Size(29, 23);
            this.btnSelectForecastFile.TabIndex = 5;
            this.btnSelectForecastFile.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectForecastFile, "Open the window to select a file");
            this.btnSelectForecastFile.UseVisualStyleBackColor = true;
            this.btnSelectForecastFile.Click += new System.EventHandler(this.btnSelectForecastFile_Click);
            // 
            // btnSelectFileSuperDettagli
            // 
            this.btnSelectFileSuperDettagli.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFileSuperDettagli.Location = new System.Drawing.Point(1071, 94);
            this.btnSelectFileSuperDettagli.Name = "btnSelectFileSuperDettagli";
            this.btnSelectFileSuperDettagli.Size = new System.Drawing.Size(29, 23);
            this.btnSelectFileSuperDettagli.TabIndex = 15;
            this.btnSelectFileSuperDettagli.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectFileSuperDettagli, "Open the window to select a file");
            this.btnSelectFileSuperDettagli.UseVisualStyleBackColor = true;
            this.btnSelectFileSuperDettagli.Click += new System.EventHandler(this.btnSelectFileSuperDettagli_Click);
            // 
            // btnSelectFileRunRate
            // 
            this.btnSelectFileRunRate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFileRunRate.Location = new System.Drawing.Point(1071, 67);
            this.btnSelectFileRunRate.Name = "btnSelectFileRunRate";
            this.btnSelectFileRunRate.Size = new System.Drawing.Size(29, 23);
            this.btnSelectFileRunRate.TabIndex = 9;
            this.btnSelectFileRunRate.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectFileRunRate, "Open the window to select a file");
            this.btnSelectFileRunRate.UseVisualStyleBackColor = true;
            this.btnSelectFileRunRate.Click += new System.EventHandler(this.btnSelectFileRunRate_Click);
            // 
            // btnOpenFileRunRate
            // 
            this.btnOpenFileRunRate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileRunRate.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileRunRate.Image")));
            this.btnOpenFileRunRate.Location = new System.Drawing.Point(1141, 66);
            this.btnOpenFileRunRate.Name = "btnOpenFileRunRate";
            this.btnOpenFileRunRate.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileRunRate.TabIndex = 11;
            this.toolTipDefault.SetToolTip(this.btnOpenFileRunRate, "Open the file with Excel");
            this.btnOpenFileRunRate.UseVisualStyleBackColor = true;
            this.btnOpenFileRunRate.Click += new System.EventHandler(this.btnOpenFileRunRate_Click);
            // 
            // btnOpenFileRunRateFolder
            // 
            this.btnOpenFileRunRateFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileRunRateFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileRunRateFolder.Image")));
            this.btnOpenFileRunRateFolder.Location = new System.Drawing.Point(1105, 66);
            this.btnOpenFileRunRateFolder.Name = "btnOpenFileRunRateFolder";
            this.btnOpenFileRunRateFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileRunRateFolder.TabIndex = 10;
            this.toolTipDefault.SetToolTip(this.btnOpenFileRunRateFolder, "Open the folder where the selected file is located");
            this.btnOpenFileRunRateFolder.UseVisualStyleBackColor = true;
            this.btnOpenFileRunRateFolder.Click += new System.EventHandler(this.btnOpenFileRunRateFolder_Click);
            // 
            // btnOpenFileSuperDettagli
            // 
            this.btnOpenFileSuperDettagli.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileSuperDettagli.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileSuperDettagli.Image")));
            this.btnOpenFileSuperDettagli.Location = new System.Drawing.Point(1141, 93);
            this.btnOpenFileSuperDettagli.Name = "btnOpenFileSuperDettagli";
            this.btnOpenFileSuperDettagli.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileSuperDettagli.TabIndex = 17;
            this.toolTipDefault.SetToolTip(this.btnOpenFileSuperDettagli, "Open the file with Excel");
            this.btnOpenFileSuperDettagli.UseVisualStyleBackColor = true;
            this.btnOpenFileSuperDettagli.Click += new System.EventHandler(this.btnOpenFileSuperDettagli_Click);
            // 
            // btnOpenFileSuperDettagliFolder
            // 
            this.btnOpenFileSuperDettagliFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileSuperDettagliFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileSuperDettagliFolder.Image")));
            this.btnOpenFileSuperDettagliFolder.Location = new System.Drawing.Point(1105, 94);
            this.btnOpenFileSuperDettagliFolder.Name = "btnOpenFileSuperDettagliFolder";
            this.btnOpenFileSuperDettagliFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileSuperDettagliFolder.TabIndex = 16;
            this.toolTipDefault.SetToolTip(this.btnOpenFileSuperDettagliFolder, "Open the folder where the selected file is located");
            this.btnOpenFileSuperDettagliFolder.UseVisualStyleBackColor = true;
            this.btnOpenFileSuperDettagliFolder.Click += new System.EventHandler(this.btnOpenFileSuperDettagliFolder_Click);
            // 
            // btnOpenFileForecast
            // 
            this.btnOpenFileForecast.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileForecast.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileForecast.Image")));
            this.btnOpenFileForecast.Location = new System.Drawing.Point(1141, 40);
            this.btnOpenFileForecast.Name = "btnOpenFileForecast";
            this.btnOpenFileForecast.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileForecast.TabIndex = 7;
            this.toolTipDefault.SetToolTip(this.btnOpenFileForecast, "Open the file with Excel");
            this.btnOpenFileForecast.UseVisualStyleBackColor = true;
            this.btnOpenFileForecast.Click += new System.EventHandler(this.btnOpenFileForecast_Click);
            // 
            // btnOpenFileForecastFolder
            // 
            this.btnOpenFileForecastFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileForecastFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileForecastFolder.Image")));
            this.btnOpenFileForecastFolder.Location = new System.Drawing.Point(1105, 40);
            this.btnOpenFileForecastFolder.Name = "btnOpenFileForecastFolder";
            this.btnOpenFileForecastFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileForecastFolder.TabIndex = 6;
            this.toolTipDefault.SetToolTip(this.btnOpenFileForecastFolder, "Open the folder where the selected file is located");
            this.btnOpenFileForecastFolder.UseVisualStyleBackColor = true;
            this.btnOpenFileForecastFolder.Click += new System.EventHandler(this.btnOpenFileForecastFolder_Click);
            // 
            // btnOpenFileBudget
            // 
            this.btnOpenFileBudget.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileBudget.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileBudget.Image")));
            this.btnOpenFileBudget.Location = new System.Drawing.Point(1141, 14);
            this.btnOpenFileBudget.Name = "btnOpenFileBudget";
            this.btnOpenFileBudget.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileBudget.TabIndex = 3;
            this.toolTipDefault.SetToolTip(this.btnOpenFileBudget, "Open the file with Excel");
            this.btnOpenFileBudget.UseVisualStyleBackColor = true;
            this.btnOpenFileBudget.Click += new System.EventHandler(this.btnOpenFileBudget_Click);
            // 
            // btnOpenDestFolder
            // 
            this.btnOpenDestFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenDestFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenDestFolder.Image")));
            this.btnOpenDestFolder.Location = new System.Drawing.Point(1105, 149);
            this.btnOpenDestFolder.Name = "btnOpenDestFolder";
            this.btnOpenDestFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenDestFolder.TabIndex = 27;
            this.toolTipDefault.SetToolTip(this.btnOpenDestFolder, "Open the destination folder");
            this.btnOpenDestFolder.UseVisualStyleBackColor = true;
            this.btnOpenDestFolder.Click += new System.EventHandler(this.btnOpenDestFolder_Click);
            // 
            // btnOpenFileBudgetFolder
            // 
            this.btnOpenFileBudgetFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileBudgetFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileBudgetFolder.Image")));
            this.btnOpenFileBudgetFolder.Location = new System.Drawing.Point(1105, 14);
            this.btnOpenFileBudgetFolder.Name = "btnOpenFileBudgetFolder";
            this.btnOpenFileBudgetFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileBudgetFolder.TabIndex = 2;
            this.toolTipDefault.SetToolTip(this.btnOpenFileBudgetFolder, "Open the folder where the selected file is located");
            this.btnOpenFileBudgetFolder.UseVisualStyleBackColor = true;
            this.btnOpenFileBudgetFolder.Click += new System.EventHandler(this.btnOpenFileBudgetFolder_Click);
            // 
            // btnBuildPresentation
            // 
            this.btnBuildPresentation.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBuildPresentation.Location = new System.Drawing.Point(1064, 575);
            this.btnBuildPresentation.Name = "btnBuildPresentation";
            this.btnBuildPresentation.Size = new System.Drawing.Size(107, 36);
            this.btnBuildPresentation.TabIndex = 33;
            this.btnBuildPresentation.Text = "&Import data and build presentation";
            this.toolTipDefault.SetToolTip(this.btnBuildPresentation, "Start building the presentation");
            this.btnBuildPresentation.UseVisualStyleBackColor = true;
            this.btnBuildPresentation.Click += new System.EventHandler(this.btnBuildPresentation_Click);
            // 
            // btnValidaInput
            // 
            this.btnValidaInput.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnValidaInput.Location = new System.Drawing.Point(1055, 215);
            this.btnValidaInput.Name = "btnValidaInput";
            this.btnValidaInput.Size = new System.Drawing.Size(120, 30);
            this.btnValidaInput.TabIndex = 28;
            this.btnValidaInput.Text = "&Validate input files";
            this.toolTipDefault.SetToolTip(this.btnValidaInput, "Validate input files and load filters information");
            this.btnValidaInput.UseVisualStyleBackColor = true;
            this.btnValidaInput.Click += new System.EventHandler(this.btnValidaInput_Click);
            // 
            // btnOpenFileCN43N
            // 
            this.btnOpenFileCN43N.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileCN43N.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileCN43N.Image")));
            this.btnOpenFileCN43N.Location = new System.Drawing.Point(1141, 119);
            this.btnOpenFileCN43N.Name = "btnOpenFileCN43N";
            this.btnOpenFileCN43N.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileCN43N.TabIndex = 24;
            this.toolTipDefault.SetToolTip(this.btnOpenFileCN43N, "Open the file with Excel");
            this.btnOpenFileCN43N.UseVisualStyleBackColor = true;
            this.btnOpenFileCN43N.Click += new System.EventHandler(this.btnOpenFileCN43N_Click);
            // 
            // btnOpenFileCN43NFolder
            // 
            this.btnOpenFileCN43NFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileCN43NFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFileCN43NFolder.Image")));
            this.btnOpenFileCN43NFolder.Location = new System.Drawing.Point(1105, 119);
            this.btnOpenFileCN43NFolder.Name = "btnOpenFileCN43NFolder";
            this.btnOpenFileCN43NFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileCN43NFolder.TabIndex = 23;
            this.toolTipDefault.SetToolTip(this.btnOpenFileCN43NFolder, "Open the folder where the selected file is located");
            this.btnOpenFileCN43NFolder.UseVisualStyleBackColor = true;
            this.btnOpenFileCN43NFolder.Click += new System.EventHandler(this.btnOpenFileCN43NFolder_Click);
            // 
            // btnSelectFileCN43N
            // 
            this.btnSelectFileCN43N.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFileCN43N.Location = new System.Drawing.Point(1071, 120);
            this.btnSelectFileCN43N.Name = "btnSelectFileCN43N";
            this.btnSelectFileCN43N.Size = new System.Drawing.Size(29, 23);
            this.btnSelectFileCN43N.TabIndex = 22;
            this.btnSelectFileCN43N.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectFileCN43N, "Open the window to select a file");
            this.btnSelectFileCN43N.UseVisualStyleBackColor = true;
            this.btnSelectFileCN43N.Click += new System.EventHandler(this.btnSelectFileCN43N_Click);
            // 
            // lblDataPeriodo
            // 
            this.lblDataPeriodo.AutoSize = true;
            this.lblDataPeriodo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDataPeriodo.Location = new System.Drawing.Point(103, 26);
            this.lblDataPeriodo.Name = "lblDataPeriodo";
            this.lblDataPeriodo.Size = new System.Drawing.Size(75, 13);
            this.lblDataPeriodo.TabIndex = 29;
            this.lblDataPeriodo.Text = "99/99/2099";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(4, 23);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(40, 13);
            this.label4.TabIndex = 47;
            this.label4.Text = "Period:";
            // 
            // btnOpenCalendar
            // 
            this.btnOpenCalendar.Image = global::PptGeneratorGUI.Properties.Resources.CalendarIcon;
            this.btnOpenCalendar.Location = new System.Drawing.Point(184, 21);
            this.btnOpenCalendar.Name = "btnOpenCalendar";
            this.btnOpenCalendar.Size = new System.Drawing.Size(29, 23);
            this.btnOpenCalendar.TabIndex = 29;
            this.btnOpenCalendar.UseVisualStyleBackColor = true;
            this.btnOpenCalendar.Click += new System.EventHandler(this.btnOpenCalendar_Click);
            // 
            // cmbFileBudgetPath
            // 
            this.cmbFileBudgetPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFileBudgetPath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbFileBudgetPath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbFileBudgetPath.FormattingEnabled = true;
            this.cmbFileBudgetPath.Location = new System.Drawing.Point(105, 16);
            this.cmbFileBudgetPath.Name = "cmbFileBudgetPath";
            this.cmbFileBudgetPath.Size = new System.Drawing.Size(960, 21);
            this.cmbFileBudgetPath.TabIndex = 0;
            this.cmbFileBudgetPath.SelectedIndexChanged += new System.EventHandler(this.cmbFileBudgetPath_SelectedIndexChanged);
            this.cmbFileBudgetPath.TextUpdate += new System.EventHandler(this.cmbFileBudgetPath_TextUpdate);
            // 
            // lblFileBudgetPath
            // 
            this.lblFileBudgetPath.AutoSize = true;
            this.lblFileBudgetPath.Location = new System.Drawing.Point(4, 19);
            this.lblFileBudgetPath.Name = "lblFileBudgetPath";
            this.lblFileBudgetPath.Size = new System.Drawing.Size(60, 13);
            this.lblFileBudgetPath.TabIndex = 6;
            this.lblFileBudgetPath.Text = "Budget file:";
            // 
            // lblResults
            // 
            this.lblResults.AutoSize = true;
            this.lblResults.Location = new System.Drawing.Point(12, 595);
            this.lblResults.Name = "lblResults";
            this.lblResults.Size = new System.Drawing.Size(45, 13);
            this.lblResults.TabIndex = 20;
            this.lblResults.Text = "Results:";
            // 
            // cmbDestinationFolderPath
            // 
            this.cmbDestinationFolderPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbDestinationFolderPath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbDestinationFolderPath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbDestinationFolderPath.FormattingEnabled = true;
            this.cmbDestinationFolderPath.Location = new System.Drawing.Point(105, 148);
            this.cmbDestinationFolderPath.Name = "cmbDestinationFolderPath";
            this.cmbDestinationFolderPath.Size = new System.Drawing.Size(960, 21);
            this.cmbDestinationFolderPath.TabIndex = 25;
            this.cmbDestinationFolderPath.SelectedIndexChanged += new System.EventHandler(this.cmbDestinationFolderPath_SelectedIndexChanged);
            this.cmbDestinationFolderPath.TextUpdate += new System.EventHandler(this.cmbDestinationFolderPath_TextUpdate);
            // 
            // wbExecutionResult
            // 
            this.wbExecutionResult.AllowWebBrowserDrop = false;
            this.wbExecutionResult.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.wbExecutionResult.CausesValidation = false;
            this.wbExecutionResult.IsWebBrowserContextMenuEnabled = false;
            this.wbExecutionResult.Location = new System.Drawing.Point(5, 616);
            this.wbExecutionResult.MinimumSize = new System.Drawing.Size(20, 20);
            this.wbExecutionResult.Name = "wbExecutionResult";
            this.wbExecutionResult.Size = new System.Drawing.Size(1165, 290);
            this.wbExecutionResult.TabIndex = 34;
            this.wbExecutionResult.WebBrowserShortcutsEnabled = false;
            this.wbExecutionResult.Navigating += new System.Windows.Forms.WebBrowserNavigatingEventHandler(this.wbExecutionResult_Navigating);
            // 
            // btnCopyOutput
            // 
            this.btnCopyOutput.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCopyOutput.AutoSize = true;
            this.btnCopyOutput.Location = new System.Drawing.Point(1007, 885);
            this.btnCopyOutput.Name = "btnCopyOutput";
            this.btnCopyOutput.Size = new System.Drawing.Size(148, 13);
            this.btnCopyOutput.TabIndex = 35;
            this.btnCopyOutput.TabStop = true;
            this.btnCopyOutput.Text = "Copy the error in the clipboard";
            this.btnCopyOutput.Visible = false;
            this.btnCopyOutput.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.btnCopyOutput_LinkClicked);
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
            this.cmbFileForecastPath.Location = new System.Drawing.Point(105, 42);
            this.cmbFileForecastPath.Name = "cmbFileForecastPath";
            this.cmbFileForecastPath.Size = new System.Drawing.Size(960, 21);
            this.cmbFileForecastPath.TabIndex = 4;
            this.cmbFileForecastPath.SelectedIndexChanged += new System.EventHandler(this.cmbFileForecastPath_SelectedIndexChanged);
            this.cmbFileForecastPath.TextUpdate += new System.EventHandler(this.cmbFileForecastPath_TextUpdate);
            // 
            // lblFileForecastPath
            // 
            this.lblFileForecastPath.AutoSize = true;
            this.lblFileForecastPath.Location = new System.Drawing.Point(4, 45);
            this.lblFileForecastPath.Name = "lblFileForecastPath";
            this.lblFileForecastPath.Size = new System.Drawing.Size(67, 13);
            this.lblFileForecastPath.TabIndex = 26;
            this.lblFileForecastPath.Text = "Forecast file:";
            // 
            // cmbFileSuperDettagliPath
            // 
            this.cmbFileSuperDettagliPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFileSuperDettagliPath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbFileSuperDettagliPath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbFileSuperDettagliPath.FormattingEnabled = true;
            this.cmbFileSuperDettagliPath.Location = new System.Drawing.Point(105, 94);
            this.cmbFileSuperDettagliPath.Name = "cmbFileSuperDettagliPath";
            this.cmbFileSuperDettagliPath.Size = new System.Drawing.Size(743, 21);
            this.cmbFileSuperDettagliPath.TabIndex = 12;
            this.cmbFileSuperDettagliPath.SelectedIndexChanged += new System.EventHandler(this.cmbFileSuperDettagliPath_SelectedIndexChanged);
            this.cmbFileSuperDettagliPath.TextUpdate += new System.EventHandler(this.cmbFileSuperDettagliPath_TextUpdate);
            // 
            // lblFileSuperDettagliPath
            // 
            this.lblFileSuperDettagliPath.AutoSize = true;
            this.lblFileSuperDettagliPath.Location = new System.Drawing.Point(4, 97);
            this.lblFileSuperDettagliPath.Name = "lblFileSuperDettagliPath";
            this.lblFileSuperDettagliPath.Size = new System.Drawing.Size(95, 13);
            this.lblFileSuperDettagliPath.TabIndex = 31;
            this.lblFileSuperDettagliPath.Text = "Super dettagli file*:";
            // 
            // cmbFileRunRatePath
            // 
            this.cmbFileRunRatePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFileRunRatePath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbFileRunRatePath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbFileRunRatePath.FormattingEnabled = true;
            this.cmbFileRunRatePath.Location = new System.Drawing.Point(105, 68);
            this.cmbFileRunRatePath.Name = "cmbFileRunRatePath";
            this.cmbFileRunRatePath.Size = new System.Drawing.Size(960, 21);
            this.cmbFileRunRatePath.TabIndex = 8;
            this.cmbFileRunRatePath.SelectedIndexChanged += new System.EventHandler(this.cmbFileRunRatePath_SelectedIndexChanged);
            this.cmbFileRunRatePath.TextUpdate += new System.EventHandler(this.cmbFileRunRatePath_TextUpdate);
            // 
            // lblFileRunRatePath
            // 
            this.lblFileRunRatePath.AutoSize = true;
            this.lblFileRunRatePath.Location = new System.Drawing.Point(4, 71);
            this.lblFileRunRatePath.Name = "lblFileRunRatePath";
            this.lblFileRunRatePath.Size = new System.Drawing.Size(71, 13);
            this.lblFileRunRatePath.TabIndex = 36;
            this.lblFileRunRatePath.Text = "Run rate file*:";
            // 
            // lblFiltriApplicabili
            // 
            this.lblFiltriApplicabili.AutoSize = true;
            this.lblFiltriApplicabili.Location = new System.Drawing.Point(4, 57);
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
            this.Table,
            this.Field,
            this.OpenFiltersSelection,
            this.SelectedValues});
            this.dgvFiltri.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dgvFiltri.Location = new System.Drawing.Point(105, 47);
            this.dgvFiltri.Name = "dgvFiltri";
            this.dgvFiltri.RowHeadersVisible = false;
            this.dgvFiltri.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvFiltri.ShowEditingIcon = false;
            this.dgvFiltri.Size = new System.Drawing.Size(1060, 268);
            this.dgvFiltri.TabIndex = 31;
            this.dgvFiltri.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvFiltri_CellContentClick);
            // 
            // Table
            // 
            this.Table.HeaderText = "Table";
            this.Table.Name = "Table";
            this.Table.ReadOnly = true;
            this.Table.Width = 64;
            // 
            // Field
            // 
            this.Field.HeaderText = "Field";
            this.Field.Name = "Field";
            this.Field.ReadOnly = true;
            this.Field.Width = 59;
            // 
            // OpenFiltersSelection
            // 
            this.OpenFiltersSelection.HeaderText = "";
            this.OpenFiltersSelection.Name = "OpenFiltersSelection";
            this.OpenFiltersSelection.ReadOnly = true;
            this.OpenFiltersSelection.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.OpenFiltersSelection.Width = 5;
            // 
            // SelectedValues
            // 
            this.SelectedValues.HeaderText = "Selected values";
            this.SelectedValues.Name = "SelectedValues";
            this.SelectedValues.ReadOnly = true;
            this.SelectedValues.Width = 123;
            // 
            // calendarPeriodo
            // 
            this.calendarPeriodo.Location = new System.Drawing.Point(0, 0);
            this.calendarPeriodo.Name = "calendarPeriodo";
            this.calendarPeriodo.TabIndex = 30;
            this.calendarPeriodo.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.calendarPeriodo_DateSelected);
            // 
            // pnlCalendar
            // 
            this.pnlCalendar.Controls.Add(this.calendarPeriodo);
            this.pnlCalendar.Location = new System.Drawing.Point(219, 23);
            this.pnlCalendar.Name = "pnlCalendar";
            this.pnlCalendar.Size = new System.Drawing.Size(226, 161);
            this.pnlCalendar.TabIndex = 49;
            this.pnlCalendar.Visible = false;
            // 
            // gbPaths
            // 
            this.gbPaths.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbPaths.Controls.Add(this.gbCN43N);
            this.gbPaths.Controls.Add(this.btnGetWbsList);
            this.gbPaths.Controls.Add(this.gbSuperDettagli);
            this.gbPaths.Controls.Add(this.btnOpenFileCN43N);
            this.gbPaths.Controls.Add(this.btnOpenFileCN43NFolder);
            this.gbPaths.Controls.Add(this.cmbFileCN43NPath);
            this.gbPaths.Controls.Add(this.btnSelectFileCN43N);
            this.gbPaths.Controls.Add(this.lblFileCN43NPath);
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
            this.gbPaths.Location = new System.Drawing.Point(4, 27);
            this.gbPaths.Name = "gbPaths";
            this.gbPaths.Size = new System.Drawing.Size(1169, 182);
            this.gbPaths.TabIndex = 50;
            this.gbPaths.TabStop = false;
            this.gbPaths.Text = "Paths";
            // 
            // gbCN43N
            // 
            this.gbCN43N.Controls.Add(this.rbCN43N_OverwriteAll);
            this.gbCN43N.Controls.Add(this.rbCN43N_Append);
            this.gbCN43N.Location = new System.Drawing.Point(799, 121);
            this.gbCN43N.Name = "gbCN43N";
            this.gbCN43N.Size = new System.Drawing.Size(266, 22);
            this.gbCN43N.TabIndex = 20;
            this.gbCN43N.TabStop = false;
            // 
            // rbCN43N_OverwriteAll
            // 
            this.rbCN43N_OverwriteAll.AutoSize = true;
            this.rbCN43N_OverwriteAll.Location = new System.Drawing.Point(183, 2);
            this.rbCN43N_OverwriteAll.Name = "rbCN43N_OverwriteAll";
            this.rbCN43N_OverwriteAll.Size = new System.Drawing.Size(83, 17);
            this.rbCN43N_OverwriteAll.TabIndex = 21;
            this.rbCN43N_OverwriteAll.Text = "Overwrite all";
            this.rbCN43N_OverwriteAll.UseVisualStyleBackColor = true;
            // 
            // rbCN43N_Append
            // 
            this.rbCN43N_Append.AutoSize = true;
            this.rbCN43N_Append.Checked = true;
            this.rbCN43N_Append.Location = new System.Drawing.Point(6, 2);
            this.rbCN43N_Append.Name = "rbCN43N_Append";
            this.rbCN43N_Append.Size = new System.Drawing.Size(170, 17);
            this.rbCN43N_Append.TabIndex = 20;
            this.rbCN43N_Append.TabStop = true;
            this.rbCN43N_Append.Text = "Append and update duplicates";
            this.rbCN43N_Append.UseVisualStyleBackColor = true;
            // 
            // btnGetWbsList
            // 
            this.btnGetWbsList.Location = new System.Drawing.Point(105, 121);
            this.btnGetWbsList.Name = "btnGetWbsList";
            this.btnGetWbsList.Size = new System.Drawing.Size(87, 21);
            this.btnGetWbsList.TabIndex = 18;
            this.btnGetWbsList.Text = "Get WBS list";
            this.btnGetWbsList.UseVisualStyleBackColor = true;
            this.btnGetWbsList.Click += new System.EventHandler(this.btnGetWbsList_Click);
            // 
            // gbSuperDettagli
            // 
            this.gbSuperDettagli.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.gbSuperDettagli.Controls.Add(this.rbSuperDettagli_Add);
            this.gbSuperDettagli.Controls.Add(this.rbSuperDettagli_ReplaceCurrentYear);
            this.gbSuperDettagli.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.gbSuperDettagli.Location = new System.Drawing.Point(869, 92);
            this.gbSuperDettagli.Margin = new System.Windows.Forms.Padding(0);
            this.gbSuperDettagli.Name = "gbSuperDettagli";
            this.gbSuperDettagli.Size = new System.Drawing.Size(196, 26);
            this.gbSuperDettagli.TabIndex = 13;
            this.gbSuperDettagli.TabStop = false;
            // 
            // rbSuperDettagli_Add
            // 
            this.rbSuperDettagli_Add.AutoSize = true;
            this.rbSuperDettagli_Add.Location = new System.Drawing.Point(146, 4);
            this.rbSuperDettagli_Add.Name = "rbSuperDettagli_Add";
            this.rbSuperDettagli_Add.Size = new System.Drawing.Size(44, 17);
            this.rbSuperDettagli_Add.TabIndex = 14;
            this.rbSuperDettagli_Add.Text = "Add";
            this.rbSuperDettagli_Add.UseVisualStyleBackColor = true;
            // 
            // rbSuperDettagli_ReplaceCurrentYear
            // 
            this.rbSuperDettagli_ReplaceCurrentYear.AutoSize = true;
            this.rbSuperDettagli_ReplaceCurrentYear.Checked = true;
            this.rbSuperDettagli_ReplaceCurrentYear.Location = new System.Drawing.Point(6, 4);
            this.rbSuperDettagli_ReplaceCurrentYear.Name = "rbSuperDettagli_ReplaceCurrentYear";
            this.rbSuperDettagli_ReplaceCurrentYear.Size = new System.Drawing.Size(124, 17);
            this.rbSuperDettagli_ReplaceCurrentYear.TabIndex = 13;
            this.rbSuperDettagli_ReplaceCurrentYear.TabStop = true;
            this.rbSuperDettagli_ReplaceCurrentYear.Text = "Replace current year";
            this.rbSuperDettagli_ReplaceCurrentYear.UseVisualStyleBackColor = true;
            // 
            // cmbFileCN43NPath
            // 
            this.cmbFileCN43NPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFileCN43NPath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbFileCN43NPath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbFileCN43NPath.FormattingEnabled = true;
            this.cmbFileCN43NPath.Location = new System.Drawing.Point(198, 121);
            this.cmbFileCN43NPath.Name = "cmbFileCN43NPath";
            this.cmbFileCN43NPath.Size = new System.Drawing.Size(595, 21);
            this.cmbFileCN43NPath.TabIndex = 19;
            this.cmbFileCN43NPath.SelectedIndexChanged += new System.EventHandler(this.cmbFileCN43NPath_SelectedIndexChanged);
            this.cmbFileCN43NPath.TextUpdate += new System.EventHandler(this.cmbFileCN43NPath_TextUpdate);
            // 
            // lblFileCN43NPath
            // 
            this.lblFileCN43NPath.AutoSize = true;
            this.lblFileCN43NPath.Location = new System.Drawing.Point(4, 124);
            this.lblFileCN43NPath.Name = "lblFileCN43NPath";
            this.lblFileCN43NPath.Size = new System.Drawing.Size(61, 13);
            this.lblFileCN43NPath.TabIndex = 41;
            this.lblFileCN43NPath.Text = "CN43N file:";
            // 
            // gbOptions
            // 
            this.gbOptions.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbOptions.Controls.Add(this.lblDataPeriodo);
            this.gbOptions.Controls.Add(this.lblFiltriApplicabili);
            this.gbOptions.Controls.Add(this.pnlCalendar);
            this.gbOptions.Controls.Add(this.label4);
            this.gbOptions.Controls.Add(this.btnOpenCalendar);
            this.gbOptions.Controls.Add(this.dgvFiltri);
            this.gbOptions.Location = new System.Drawing.Point(5, 251);
            this.gbOptions.Name = "gbOptions";
            this.gbOptions.Size = new System.Drawing.Size(1169, 321);
            this.gbOptions.TabIndex = 51;
            this.gbOptions.TabStop = false;
            this.gbOptions.Text = "Options";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1,
            this.sessionToolStripMenuItem,
            this.updatePresentationToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1185, 24);
            this.menuStrip1.TabIndex = 53;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openDataSouceExcelFileToolStripMenuItem,
            this.openSouceFilesFolderToolStripMenuItem});
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(79, 20);
            this.toolStripMenuItem1.Text = "Data&Source";
            // 
            // openDataSouceExcelFileToolStripMenuItem
            // 
            this.openDataSouceExcelFileToolStripMenuItem.Name = "openDataSouceExcelFileToolStripMenuItem";
            this.openDataSouceExcelFileToolStripMenuItem.Size = new System.Drawing.Size(210, 22);
            this.openDataSouceExcelFileToolStripMenuItem.Text = "Open &DataSouce Excel file";
            this.openDataSouceExcelFileToolStripMenuItem.Click += new System.EventHandler(this.openDataSouceExcelFileToolStripMenuItem_Click);
            // 
            // openSouceFilesFolderToolStripMenuItem
            // 
            this.openSouceFilesFolderToolStripMenuItem.Name = "openSouceFilesFolderToolStripMenuItem";
            this.openSouceFilesFolderToolStripMenuItem.Size = new System.Drawing.Size(210, 22);
            this.openSouceFilesFolderToolStripMenuItem.Text = "Open DataSource &Folder";
            this.openSouceFilesFolderToolStripMenuItem.Click += new System.EventHandler(this.openSouceFilesFolderToolStripMenuItem_Click);
            // 
            // sessionToolStripMenuItem
            // 
            this.sessionToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.loadLastSessionPathsToolStripMenuItem,
            this.cleanCurrentsessionToolStripMenuItem,
            this.deletePathsHistoryToolStripMenuItem});
            this.sessionToolStripMenuItem.Name = "sessionToolStripMenuItem";
            this.sessionToolStripMenuItem.Size = new System.Drawing.Size(58, 20);
            this.sessionToolStripMenuItem.Text = "&Session";
            // 
            // loadLastSessionPathsToolStripMenuItem
            // 
            this.loadLastSessionPathsToolStripMenuItem.Name = "loadLastSessionPathsToolStripMenuItem";
            this.loadLastSessionPathsToolStripMenuItem.Size = new System.Drawing.Size(194, 22);
            this.loadLastSessionPathsToolStripMenuItem.Text = "&Load last session paths";
            this.loadLastSessionPathsToolStripMenuItem.Click += new System.EventHandler(this.loadLastSessionPathsToolStripMenuItem1_Click);
            // 
            // cleanCurrentsessionToolStripMenuItem
            // 
            this.cleanCurrentsessionToolStripMenuItem.Name = "cleanCurrentsessionToolStripMenuItem";
            this.cleanCurrentsessionToolStripMenuItem.Size = new System.Drawing.Size(194, 22);
            this.cleanCurrentsessionToolStripMenuItem.Text = "&Clean current session";
            this.cleanCurrentsessionToolStripMenuItem.Click += new System.EventHandler(this.cleanCurrentsessionToolStripMenuItem_Click);
            // 
            // deletePathsHistoryToolStripMenuItem
            // 
            this.deletePathsHistoryToolStripMenuItem.Name = "deletePathsHistoryToolStripMenuItem";
            this.deletePathsHistoryToolStripMenuItem.Size = new System.Drawing.Size(194, 22);
            this.deletePathsHistoryToolStripMenuItem.Text = "&Delete paths history";
            this.deletePathsHistoryToolStripMenuItem.Click += new System.EventHandler(this.deletePathsHistoryToolStripMenuItem_Click);
            // 
            // updatePresentationToolStripMenuItem
            // 
            this.updatePresentationToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.updatePresentationToolStripMenuItem1});
            this.updatePresentationToolStripMenuItem.Name = "updatePresentationToolStripMenuItem";
            this.updatePresentationToolStripMenuItem.Size = new System.Drawing.Size(85, 20);
            this.updatePresentationToolStripMenuItem.Text = "&Presentation";
            // 
            // updatePresentationToolStripMenuItem1
            // 
            this.updatePresentationToolStripMenuItem1.Name = "updatePresentationToolStripMenuItem1";
            this.updatePresentationToolStripMenuItem1.Size = new System.Drawing.Size(181, 22);
            this.updatePresentationToolStripMenuItem1.Text = "Update &Presentation";
            this.updatePresentationToolStripMenuItem1.Click += new System.EventHandler(this.updatePresentationToolStripMenuItem_Click);
            // 
            // lblElaborazioneInCorso
            // 
            this.lblElaborazioneInCorso.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblElaborazioneInCorso.BackColor = System.Drawing.SystemColors.ControlDark;
            this.lblElaborazioneInCorso.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblElaborazioneInCorso.Font = new System.Drawing.Font("Microsoft Sans Serif", 30F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblElaborazioneInCorso.Location = new System.Drawing.Point(63, 27);
            this.lblElaborazioneInCorso.Name = "lblElaborazioneInCorso";
            this.lblElaborazioneInCorso.Size = new System.Drawing.Size(1065, 879);
            this.lblElaborazioneInCorso.TabIndex = 54;
            this.lblElaborazioneInCorso.Text = "Working in progress...this might take several minutes.";
            this.lblElaborazioneInCorso.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblElaborazioneInCorso.Visible = false;
            // 
            // btnTryBuildPresentationOnly
            // 
            this.btnTryBuildPresentationOnly.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnTryBuildPresentationOnly.Location = new System.Drawing.Point(948, 575);
            this.btnTryBuildPresentationOnly.Margin = new System.Windows.Forms.Padding(2);
            this.btnTryBuildPresentationOnly.Name = "btnTryBuildPresentationOnly";
            this.btnTryBuildPresentationOnly.Size = new System.Drawing.Size(107, 36);
            this.btnTryBuildPresentationOnly.TabIndex = 32;
            this.btnTryBuildPresentationOnly.Text = "Build presentation ONLY";
            this.btnTryBuildPresentationOnly.UseVisualStyleBackColor = true;
            this.btnTryBuildPresentationOnly.Visible = false;
            this.btnTryBuildPresentationOnly.Click += new System.EventHandler(this.btnTryBuildPresentationOnly_Click);
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
            this.ClientSize = new System.Drawing.Size(1185, 924);
            this.Controls.Add(this.lblElaborazioneInCorso);
            this.Controls.Add(this.btnTryBuildPresentationOnly);
            this.Controls.Add(this.btnValidaInput);
            this.Controls.Add(this.btnBuildPresentation);
            this.Controls.Add(this.btnCopyOutput);
            this.Controls.Add(this.lblResults);
            this.Controls.Add(this.statusStrip);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.gbPaths);
            this.Controls.Add(this.gbOptions);
            this.Controls.Add(this.wbExecutionResult);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(1000, 750);
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PowerPoint Generator";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.statusStrip.ResumeLayout(false);
            this.statusStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvFiltri)).EndInit();
            this.pnlCalendar.ResumeLayout(false);
            this.gbPaths.ResumeLayout(false);
            this.gbPaths.PerformLayout();
            this.gbCN43N.ResumeLayout(false);
            this.gbCN43N.PerformLayout();
            this.gbSuperDettagli.ResumeLayout(false);
            this.gbSuperDettagli.PerformLayout();
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
        private System.Windows.Forms.Label lblResults;
        private System.Windows.Forms.ComboBox cmbDestinationFolderPath;
        private System.Windows.Forms.Button btnSelectDestinationFolder;
        private System.Windows.Forms.Label lblCartellaOutputPath;
        private System.Windows.Forms.Button btnOpenFileBudgetFolder;
        private System.Windows.Forms.Button btnOpenDestFolder;
        private System.Windows.Forms.Button btnOpenFileBudget;
        private System.Windows.Forms.WebBrowser wbExecutionResult;
        private System.Windows.Forms.LinkLabel btnCopyOutput;
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
        private System.Windows.Forms.Label lblDataPeriodo;
        private System.Windows.Forms.Button btnOpenCalendar;
        private System.Windows.Forms.MonthCalendar calendarPeriodo;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel pnlCalendar;
        private System.Windows.Forms.GroupBox gbPaths;
        private System.Windows.Forms.GroupBox gbOptions;
        private System.Windows.Forms.Button btnValidaInput;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem openSouceFilesFolderToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openDataSouceExcelFileToolStripMenuItem;
        private System.Windows.Forms.DataGridViewTextBoxColumn Table;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field;
        private System.Windows.Forms.DataGridViewButtonColumn OpenFiltersSelection;
        private System.Windows.Forms.DataGridViewTextBoxColumn SelectedValues;
        private System.Windows.Forms.Label lblElaborazioneInCorso;
        private System.Windows.Forms.ToolStripMenuItem sessionToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem loadLastSessionPathsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cleanCurrentsessionToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem deletePathsHistoryToolStripMenuItem;
        private System.Windows.Forms.Button btnOpenFileCN43N;
        private System.Windows.Forms.Button btnOpenFileCN43NFolder;
        private System.Windows.Forms.ComboBox cmbFileCN43NPath;
        private System.Windows.Forms.Button btnSelectFileCN43N;
        private System.Windows.Forms.Label lblFileCN43NPath;
        private System.Windows.Forms.Button btnTryBuildPresentationOnly;
        private System.Windows.Forms.ToolStripMenuItem updatePresentationToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem updatePresentationToolStripMenuItem1;
        private System.Windows.Forms.RadioButton rbSuperDettagli_Add;
        private System.Windows.Forms.RadioButton rbSuperDettagli_ReplaceCurrentYear;
        private System.Windows.Forms.GroupBox gbSuperDettagli;
        private System.Windows.Forms.Button btnGetWbsList;
        private System.Windows.Forms.GroupBox gbCN43N;
        private System.Windows.Forms.RadioButton rbCN43N_OverwriteAll;
        private System.Windows.Forms.RadioButton rbCN43N_Append;
    }
}

