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
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripProgressBar = new System.Windows.Forms.ToolStripProgressBar();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.txtStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.lblVersion = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolTipDefault = new System.Windows.Forms.ToolTip(this.components);
            this.btnSelectFileBudget = new System.Windows.Forms.Button();
            this.btnSelectDestinationFolder = new System.Windows.Forms.Button();
            this.lblCartellaOutputPath = new System.Windows.Forms.Label();
            this.btnOpenFileBudget = new System.Windows.Forms.Button();
            this.btnOpenDestFolder = new System.Windows.Forms.Button();
            this.btnOpenFileBudgetFolder = new System.Windows.Forms.Button();
            this.btnOpenFileForecast = new System.Windows.Forms.Button();
            this.btnOpenFileForecastFolder = new System.Windows.Forms.Button();
            this.btnSelectForecastFile = new System.Windows.Forms.Button();
            this.btnOpenFileSuperDettagli = new System.Windows.Forms.Button();
            this.btnOpenFileSuperDettagliFolder = new System.Windows.Forms.Button();
            this.btnSelectFileSuperDettagli = new System.Windows.Forms.Button();
            this.btnOpenFileRanRate = new System.Windows.Forms.Button();
            this.btnOpenFileRanRateFolder = new System.Windows.Forms.Button();
            this.btnSelectFileRanRate = new System.Windows.Forms.Button();
            this.cmbFileBudgetPath = new System.Windows.Forms.ComboBox();
            this.lblFileBudgetPath = new System.Windows.Forms.Label();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.label2 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItemClear = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItemOpenConfigFolder = new System.Windows.Forms.ToolStripMenuItem();
            this.cmbDestinationFolderPath = new System.Windows.Forms.ComboBox();
            this.wbExecutionResult = new System.Windows.Forms.WebBrowser();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.updateReportsBackgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.btnClear = new System.Windows.Forms.LinkLabel();
            this.btnCopyError = new System.Windows.Forms.LinkLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.cmbFileForecastPath = new System.Windows.Forms.ComboBox();
            this.lblFileForecastPath = new System.Windows.Forms.Label();
            this.cmbFileSuperDettagliPath = new System.Windows.Forms.ComboBox();
            this.lblFileSuperDettagliPath = new System.Windows.Forms.Label();
            this.cmbFileRanRatePath = new System.Windows.Forms.ComboBox();
            this.lblFileRanRatePath = new System.Windows.Forms.Label();
            this.btnCreaPresentazione = new System.Windows.Forms.Button();
            this.bfbDestFolder = new WK.Libraries.BetterFolderBrowserNS.BetterFolderBrowser(this.components);
            this.statusStrip1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // statusStrip1
            // 
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripProgressBar,
            this.toolStripStatusLabel1,
            this.txtStatusLabel,
            this.lblVersion});
            this.statusStrip1.Location = new System.Drawing.Point(0, 673);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new System.Windows.Forms.Padding(2, 0, 14, 0);
            this.statusStrip1.Size = new System.Drawing.Size(1093, 22);
            this.statusStrip1.TabIndex = 21;
            this.statusStrip1.Text = "statusStrip1";
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
            this.lblVersion.Size = new System.Drawing.Size(985, 17);
            this.lblVersion.Spring = true;
            this.lblVersion.Text = "[VERSIONE]";
            this.lblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnSelectFileBudget
            // 
            this.btnSelectFileBudget.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFileBudget.Location = new System.Drawing.Point(982, 30);
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
            this.btnSelectDestinationFolder.Location = new System.Drawing.Point(982, 137);
            this.btnSelectDestinationFolder.Name = "btnSelectDestinationFolder";
            this.btnSelectDestinationFolder.Size = new System.Drawing.Size(29, 23);
            this.btnSelectDestinationFolder.TabIndex = 5;
            this.btnSelectDestinationFolder.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectDestinationFolder, "Apre la finestra di selezione cartella");
            this.btnSelectDestinationFolder.UseVisualStyleBackColor = true;
            this.btnSelectDestinationFolder.Click += new System.EventHandler(this.btnSelectDestinationFolder_Click);
            // 
            // lblCartellaOutputPath
            // 
            this.lblCartellaOutputPath.AutoSize = true;
            this.lblCartellaOutputPath.Location = new System.Drawing.Point(6, 142);
            this.lblCartellaOutputPath.Name = "lblCartellaOutputPath";
            this.lblCartellaOutputPath.Size = new System.Drawing.Size(118, 13);
            this.lblCartellaOutputPath.TabIndex = 11;
            this.lblCartellaOutputPath.Text = "Cartella di destinazione:";
            this.toolTipDefault.SetToolTip(this.lblCartellaOutputPath, "Cartella nella quale vettanno salvati i file di output");
            // 
            // btnOpenFileBudget
            // 
            this.btnOpenFileBudget.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileBudget.Image = global::PptGeneratorGUI.Properties.Resources.png_clipart_spreadsheet_computer_icons_google_docs_microsoft_excel_table_angle_furniture_small1;
            this.btnOpenFileBudget.Location = new System.Drawing.Point(1055, 30);
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
            this.btnOpenDestFolder.Image = global::PptGeneratorGUI.Properties.Resources.free_yellow_open_folder_icon_11567_thumb_Small;
            this.btnOpenDestFolder.Location = new System.Drawing.Point(1019, 137);
            this.btnOpenDestFolder.Name = "btnOpenDestFolder";
            this.btnOpenDestFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenDestFolder.TabIndex = 6;
            this.toolTipDefault.SetToolTip(this.btnOpenDestFolder, "Apre la cartella di destinazione");
            this.btnOpenDestFolder.UseVisualStyleBackColor = true;
            this.btnOpenDestFolder.Click += new System.EventHandler(this.btnOpenDestFolder_Click);
            // 
            // btnOpenFileBudgetFolder
            // 
            this.btnOpenFileBudgetFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileBudgetFolder.Image = global::PptGeneratorGUI.Properties.Resources.free_yellow_open_folder_icon_11567_thumb_Small;
            this.btnOpenFileBudgetFolder.Location = new System.Drawing.Point(1019, 30);
            this.btnOpenFileBudgetFolder.Name = "btnOpenFileBudgetFolder";
            this.btnOpenFileBudgetFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileBudgetFolder.TabIndex = 2;
            this.toolTipDefault.SetToolTip(this.btnOpenFileBudgetFolder, "Apre la cartella contenente il file report");
            this.btnOpenFileBudgetFolder.UseVisualStyleBackColor = true;
            this.btnOpenFileBudgetFolder.Click += new System.EventHandler(this.btnOpenFileBudgetFolder_Click);
            // 
            // btnOpenFileForecast
            // 
            this.btnOpenFileForecast.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileForecast.Image = global::PptGeneratorGUI.Properties.Resources.png_clipart_spreadsheet_computer_icons_google_docs_microsoft_excel_table_angle_furniture_small1;
            this.btnOpenFileForecast.Location = new System.Drawing.Point(1055, 56);
            this.btnOpenFileForecast.Name = "btnOpenFileForecast";
            this.btnOpenFileForecast.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileForecast.TabIndex = 25;
            this.toolTipDefault.SetToolTip(this.btnOpenFileForecast, "Apre il file report");
            this.btnOpenFileForecast.UseVisualStyleBackColor = true;
            this.btnOpenFileForecast.Click += new System.EventHandler(this.btnOpenFileForecast_Click);
            // 
            // btnOpenFileForecastFolder
            // 
            this.btnOpenFileForecastFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileForecastFolder.Image = global::PptGeneratorGUI.Properties.Resources.free_yellow_open_folder_icon_11567_thumb_Small;
            this.btnOpenFileForecastFolder.Location = new System.Drawing.Point(1019, 56);
            this.btnOpenFileForecastFolder.Name = "btnOpenFileForecastFolder";
            this.btnOpenFileForecastFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileForecastFolder.TabIndex = 24;
            this.toolTipDefault.SetToolTip(this.btnOpenFileForecastFolder, "Apre la cartella contenente il file report");
            this.btnOpenFileForecastFolder.UseVisualStyleBackColor = true;
            this.btnOpenFileForecastFolder.Click += new System.EventHandler(this.btnOpenFileForecastFolder_Click);
            // 
            // btnSelectForecastFile
            // 
            this.btnSelectForecastFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectForecastFile.Location = new System.Drawing.Point(982, 56);
            this.btnSelectForecastFile.Name = "btnSelectForecastFile";
            this.btnSelectForecastFile.Size = new System.Drawing.Size(29, 23);
            this.btnSelectForecastFile.TabIndex = 23;
            this.btnSelectForecastFile.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectForecastFile, "Apre la finestra di selezione file");
            this.btnSelectForecastFile.UseVisualStyleBackColor = true;
            this.btnSelectForecastFile.Click += new System.EventHandler(this.btnSelectForecastFile_Click);
            // 
            // btnOpenFileSuperDettagli
            // 
            this.btnOpenFileSuperDettagli.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileSuperDettagli.Image = global::PptGeneratorGUI.Properties.Resources.png_clipart_spreadsheet_computer_icons_google_docs_microsoft_excel_table_angle_furniture_small1;
            this.btnOpenFileSuperDettagli.Location = new System.Drawing.Point(1055, 82);
            this.btnOpenFileSuperDettagli.Name = "btnOpenFileSuperDettagli";
            this.btnOpenFileSuperDettagli.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileSuperDettagli.TabIndex = 30;
            this.toolTipDefault.SetToolTip(this.btnOpenFileSuperDettagli, "Apre il file report");
            this.btnOpenFileSuperDettagli.UseVisualStyleBackColor = true;
            this.btnOpenFileSuperDettagli.Click += new System.EventHandler(this.btnOpenFileSuperDettagli_Click);
            // 
            // btnOpenFileSuperDettagliFolder
            // 
            this.btnOpenFileSuperDettagliFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileSuperDettagliFolder.Image = global::PptGeneratorGUI.Properties.Resources.free_yellow_open_folder_icon_11567_thumb_Small;
            this.btnOpenFileSuperDettagliFolder.Location = new System.Drawing.Point(1019, 82);
            this.btnOpenFileSuperDettagliFolder.Name = "btnOpenFileSuperDettagliFolder";
            this.btnOpenFileSuperDettagliFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileSuperDettagliFolder.TabIndex = 29;
            this.toolTipDefault.SetToolTip(this.btnOpenFileSuperDettagliFolder, "Apre la cartella contenente il file report");
            this.btnOpenFileSuperDettagliFolder.UseVisualStyleBackColor = true;
            this.btnOpenFileSuperDettagliFolder.Click += new System.EventHandler(this.btnOpenFileSuperDettagliFolder_Click);
            // 
            // btnSelectFileSuperDettagli
            // 
            this.btnSelectFileSuperDettagli.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFileSuperDettagli.Location = new System.Drawing.Point(982, 82);
            this.btnSelectFileSuperDettagli.Name = "btnSelectFileSuperDettagli";
            this.btnSelectFileSuperDettagli.Size = new System.Drawing.Size(29, 23);
            this.btnSelectFileSuperDettagli.TabIndex = 28;
            this.btnSelectFileSuperDettagli.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectFileSuperDettagli, "Apre la finestra di selezione file");
            this.btnSelectFileSuperDettagli.UseVisualStyleBackColor = true;
            this.btnSelectFileSuperDettagli.Click += new System.EventHandler(this.btnSelectFileSuperDettagli_Click);
            // 
            // btnOpenFileRanRate
            // 
            this.btnOpenFileRanRate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileRanRate.Image = global::PptGeneratorGUI.Properties.Resources.png_clipart_spreadsheet_computer_icons_google_docs_microsoft_excel_table_angle_furniture_small1;
            this.btnOpenFileRanRate.Location = new System.Drawing.Point(1055, 109);
            this.btnOpenFileRanRate.Name = "btnOpenFileRanRate";
            this.btnOpenFileRanRate.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileRanRate.TabIndex = 35;
            this.toolTipDefault.SetToolTip(this.btnOpenFileRanRate, "Apre il file report");
            this.btnOpenFileRanRate.UseVisualStyleBackColor = true;
            this.btnOpenFileRanRate.Click += new System.EventHandler(this.btnOpenFileRanRate_Click);
            // 
            // btnOpenFileRanRateFolder
            // 
            this.btnOpenFileRanRateFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFileRanRateFolder.Image = global::PptGeneratorGUI.Properties.Resources.free_yellow_open_folder_icon_11567_thumb_Small;
            this.btnOpenFileRanRateFolder.Location = new System.Drawing.Point(1019, 109);
            this.btnOpenFileRanRateFolder.Name = "btnOpenFileRanRateFolder";
            this.btnOpenFileRanRateFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFileRanRateFolder.TabIndex = 34;
            this.toolTipDefault.SetToolTip(this.btnOpenFileRanRateFolder, "Apre la cartella contenente il file report");
            this.btnOpenFileRanRateFolder.UseVisualStyleBackColor = true;
            this.btnOpenFileRanRateFolder.Click += new System.EventHandler(this.btnOpenFileRanRateFolder_Click);
            // 
            // btnSelectFileRanRate
            // 
            this.btnSelectFileRanRate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFileRanRate.Location = new System.Drawing.Point(982, 109);
            this.btnSelectFileRanRate.Name = "btnSelectFileRanRate";
            this.btnSelectFileRanRate.Size = new System.Drawing.Size(29, 23);
            this.btnSelectFileRanRate.TabIndex = 33;
            this.btnSelectFileRanRate.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectFileRanRate, "Apre la finestra di selezione file");
            this.btnSelectFileRanRate.UseVisualStyleBackColor = true;
            this.btnSelectFileRanRate.Click += new System.EventHandler(this.btnSelectFileRanRate_Click);
            // 
            // cmbFileBudgetPath
            // 
            this.cmbFileBudgetPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFileBudgetPath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbFileBudgetPath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbFileBudgetPath.FormattingEnabled = true;
            this.cmbFileBudgetPath.Location = new System.Drawing.Point(124, 32);
            this.cmbFileBudgetPath.Name = "cmbFileBudgetPath";
            this.cmbFileBudgetPath.Size = new System.Drawing.Size(851, 21);
            this.cmbFileBudgetPath.TabIndex = 0;
            this.cmbFileBudgetPath.SelectedIndexChanged += new System.EventHandler(this.cmbFileBudgetPath_SelectedIndexChanged);
            this.cmbFileBudgetPath.TextUpdate += new System.EventHandler(this.cmbFileBudgetPath_TextUpdate);
            // 
            // lblFileBudgetPath
            // 
            this.lblFileBudgetPath.AutoSize = true;
            this.lblFileBudgetPath.Location = new System.Drawing.Point(6, 35);
            this.lblFileBudgetPath.Name = "lblFileBudgetPath";
            this.lblFileBudgetPath.Size = new System.Drawing.Size(73, 13);
            this.lblFileBudgetPath.TabIndex = 6;
            this.lblFileBudgetPath.Text = "File \"Budget\":";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 186);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(132, 13);
            this.label2.TabIndex = 20;
            this.label2.Text = "Risultato dell\'aleborazione:";
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(4, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(1093, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItemClear,
            this.toolStripMenuItemOpenConfigFolder});
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(87, 20);
            this.toolStripMenuItem1.Text = "Impostazioni";
            // 
            // toolStripMenuItemClear
            // 
            this.toolStripMenuItemClear.Name = "toolStripMenuItemClear";
            this.toolStripMenuItemClear.Size = new System.Drawing.Size(256, 22);
            this.toolStripMenuItemClear.Text = "&Svuota cronologia menu a tendina";
            this.toolStripMenuItemClear.Click += new System.EventHandler(this.toolStripMenuItemClear_Click);
            // 
            // toolStripMenuItemOpenConfigFolder
            // 
            this.toolStripMenuItemOpenConfigFolder.Name = "toolStripMenuItemOpenConfigFolder";
            this.toolStripMenuItemOpenConfigFolder.Size = new System.Drawing.Size(256, 22);
            this.toolStripMenuItemOpenConfigFolder.Text = "&Apri cartella di configurazione";
            this.toolStripMenuItemOpenConfigFolder.Click += new System.EventHandler(this.toolStripMenuItemOpenConfigFolder_Click);
            // 
            // cmbDestinationFolderPath
            // 
            this.cmbDestinationFolderPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbDestinationFolderPath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbDestinationFolderPath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbDestinationFolderPath.FormattingEnabled = true;
            this.cmbDestinationFolderPath.Location = new System.Drawing.Point(124, 138);
            this.cmbDestinationFolderPath.Name = "cmbDestinationFolderPath";
            this.cmbDestinationFolderPath.Size = new System.Drawing.Size(851, 21);
            this.cmbDestinationFolderPath.TabIndex = 4;
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
            this.wbExecutionResult.Location = new System.Drawing.Point(9, 205);
            this.wbExecutionResult.MinimumSize = new System.Drawing.Size(20, 20);
            this.wbExecutionResult.Name = "wbExecutionResult";
            this.wbExecutionResult.Size = new System.Drawing.Size(1075, 465);
            this.wbExecutionResult.TabIndex = 17;
            this.wbExecutionResult.Navigating += new System.Windows.Forms.WebBrowserNavigatingEventHandler(this.wbExecutionResult_Navigating_1);
            // 
            // updateReportsBackgroundWorker
            // 
            this.updateReportsBackgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.updateReportsBackgroundWorker_DoWork);
            this.updateReportsBackgroundWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.updateReportsBackgroundWorker_RunWorkerCompleted);
            // 
            // btnClear
            // 
            this.btnClear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClear.AutoSize = true;
            this.btnClear.Location = new System.Drawing.Point(934, 214);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(123, 13);
            this.btnClear.TabIndex = 16;
            this.btnClear.TabStop = true;
            this.btnClear.Text = "Pulisci sessione corrente";
            this.btnClear.Visible = false;
            this.btnClear.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.btnClear_LinkClicked);
            // 
            // btnCopyError
            // 
            this.btnCopyError.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCopyError.AutoSize = true;
            this.btnCopyError.Location = new System.Drawing.Point(993, 648);
            this.btnCopyError.Name = "btnCopyError";
            this.btnCopyError.Size = new System.Drawing.Size(64, 13);
            this.btnCopyError.TabIndex = 17;
            this.btnCopyError.TabStop = true;
            this.btnCopyError.Text = "Copia errore";
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
            this.cmbFileForecastPath.Location = new System.Drawing.Point(124, 58);
            this.cmbFileForecastPath.Name = "cmbFileForecastPath";
            this.cmbFileForecastPath.Size = new System.Drawing.Size(851, 21);
            this.cmbFileForecastPath.TabIndex = 22;
            this.cmbFileForecastPath.SelectedIndexChanged += new System.EventHandler(this.cmbFileForecastPath_SelectedIndexChanged);
            this.cmbFileForecastPath.TextUpdate += new System.EventHandler(this.cmbFileForecastPath_TextUpdate);
            // 
            // lblFileForecastPath
            // 
            this.lblFileForecastPath.AutoSize = true;
            this.lblFileForecastPath.Location = new System.Drawing.Point(6, 61);
            this.lblFileForecastPath.Name = "lblFileForecastPath";
            this.lblFileForecastPath.Size = new System.Drawing.Size(80, 13);
            this.lblFileForecastPath.TabIndex = 26;
            this.lblFileForecastPath.Text = "File \"Forecast\":";
            // 
            // cmbFileSuperDettagliPath
            // 
            this.cmbFileSuperDettagliPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFileSuperDettagliPath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbFileSuperDettagliPath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbFileSuperDettagliPath.FormattingEnabled = true;
            this.cmbFileSuperDettagliPath.Location = new System.Drawing.Point(124, 84);
            this.cmbFileSuperDettagliPath.Name = "cmbFileSuperDettagliPath";
            this.cmbFileSuperDettagliPath.Size = new System.Drawing.Size(851, 21);
            this.cmbFileSuperDettagliPath.TabIndex = 27;
            this.cmbFileSuperDettagliPath.SelectedIndexChanged += new System.EventHandler(this.cmbFileSuperDettagliPath_SelectedIndexChanged);
            this.cmbFileSuperDettagliPath.TextUpdate += new System.EventHandler(this.cmbFileSuperDettagliPath_TextUpdate);
            // 
            // lblFileSuperDettagliPath
            // 
            this.lblFileSuperDettagliPath.AutoSize = true;
            this.lblFileSuperDettagliPath.Location = new System.Drawing.Point(6, 87);
            this.lblFileSuperDettagliPath.Name = "lblFileSuperDettagliPath";
            this.lblFileSuperDettagliPath.Size = new System.Drawing.Size(104, 13);
            this.lblFileSuperDettagliPath.TabIndex = 31;
            this.lblFileSuperDettagliPath.Text = "File \"Super dettagli\":";
            // 
            // cmbFileRanRatePath
            // 
            this.cmbFileRanRatePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFileRanRatePath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbFileRanRatePath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbFileRanRatePath.FormattingEnabled = true;
            this.cmbFileRanRatePath.Location = new System.Drawing.Point(124, 111);
            this.cmbFileRanRatePath.Name = "cmbFileRanRatePath";
            this.cmbFileRanRatePath.Size = new System.Drawing.Size(851, 21);
            this.cmbFileRanRatePath.TabIndex = 32;
            this.cmbFileRanRatePath.SelectedIndexChanged += new System.EventHandler(this.cmbFileRanRatePath_SelectedIndexChanged);
            this.cmbFileRanRatePath.TextUpdate += new System.EventHandler(this.cmbFileRanRatePath_TextUpdate);
            // 
            // lblFileRanRatePath
            // 
            this.lblFileRanRatePath.AutoSize = true;
            this.lblFileRanRatePath.Location = new System.Drawing.Point(6, 114);
            this.lblFileRanRatePath.Name = "lblFileRanRatePath";
            this.lblFileRanRatePath.Size = new System.Drawing.Size(80, 13);
            this.lblFileRanRatePath.TabIndex = 36;
            this.lblFileRanRatePath.Text = "File \"Ran rate\":";
            // 
            // btnCreaPresentazione
            // 
            this.btnCreaPresentazione.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCreaPresentazione.Location = new System.Drawing.Point(947, 166);
            this.btnCreaPresentazione.Name = "btnCreaPresentazione";
            this.btnCreaPresentazione.Size = new System.Drawing.Size(137, 33);
            this.btnCreaPresentazione.TabIndex = 37;
            this.btnCreaPresentazione.Text = "&Crea presentazione";
            this.btnCreaPresentazione.UseVisualStyleBackColor = true;
            this.btnCreaPresentazione.Click += new System.EventHandler(this.btnCreaPresentazione_Click);
            // 
            // bfbDestFolder
            // 
            this.bfbDestFolder.Multiselect = false;
            this.bfbDestFolder.RootFolder = "C:\\Users\\DELL\\Desktop";
            this.bfbDestFolder.Title = "Selezionare la cartella di destinazione del report...";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1093, 695);
            this.Controls.Add(this.btnCreaPresentazione);
            this.Controls.Add(this.btnOpenFileRanRate);
            this.Controls.Add(this.btnOpenFileRanRateFolder);
            this.Controls.Add(this.cmbFileRanRatePath);
            this.Controls.Add(this.btnSelectFileRanRate);
            this.Controls.Add(this.lblFileRanRatePath);
            this.Controls.Add(this.btnOpenFileSuperDettagli);
            this.Controls.Add(this.btnOpenFileSuperDettagliFolder);
            this.Controls.Add(this.cmbFileSuperDettagliPath);
            this.Controls.Add(this.btnSelectFileSuperDettagli);
            this.Controls.Add(this.lblFileSuperDettagliPath);
            this.Controls.Add(this.btnOpenFileForecast);
            this.Controls.Add(this.btnOpenFileForecastFolder);
            this.Controls.Add(this.cmbFileForecastPath);
            this.Controls.Add(this.btnSelectForecastFile);
            this.Controls.Add(this.lblFileForecastPath);
            this.Controls.Add(this.btnCopyError);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.wbExecutionResult);
            this.Controls.Add(this.btnOpenFileBudget);
            this.Controls.Add(this.btnOpenDestFolder);
            this.Controls.Add(this.btnOpenFileBudgetFolder);
            this.Controls.Add(this.cmbDestinationFolderPath);
            this.Controls.Add(this.btnSelectDestinationFolder);
            this.Controls.Add(this.lblCartellaOutputPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cmbFileBudgetPath);
            this.Controls.Add(this.btnSelectFileBudget);
            this.Controls.Add(this.lblFileBudgetPath);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.MinimumSize = new System.Drawing.Size(458, 358);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PowerPoint Generator";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel txtStatusLabel;
        private System.Windows.Forms.ToolTip toolTipDefault;
        private System.Windows.Forms.ComboBox cmbFileBudgetPath;
        private System.Windows.Forms.Button btnSelectFileBudget;
        private System.Windows.Forms.Label lblFileBudgetPath;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemClear;
        private System.Windows.Forms.ComboBox cmbDestinationFolderPath;
        private System.Windows.Forms.Button btnSelectDestinationFolder;
        private System.Windows.Forms.Label lblCartellaOutputPath;
        private System.Windows.Forms.Button btnOpenFileBudgetFolder;
        private System.Windows.Forms.Button btnOpenDestFolder;
        private System.Windows.Forms.Button btnOpenFileBudget;
        private System.Windows.Forms.WebBrowser wbExecutionResult;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.ComponentModel.BackgroundWorker updateReportsBackgroundWorker;
        private System.Windows.Forms.LinkLabel btnClear;
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
    }
}

