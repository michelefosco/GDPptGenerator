namespace VendorReportRefresher
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
            this.btnSelectFile1 = new System.Windows.Forms.Button();
            this.btnSelectDestinationFolder = new System.Windows.Forms.Button();
            this.lblCartellaOutputPath = new System.Windows.Forms.Label();
            this.btnOpenFile1File = new System.Windows.Forms.Button();
            this.btnOpenDestFolder = new System.Windows.Forms.Button();
            this.btnOpenFile1Folder = new System.Windows.Forms.Button();
            this.cmbFile1Path = new System.Windows.Forms.ComboBox();
            this.lblFile1Path = new System.Windows.Forms.Label();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnStart = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItemClear = new System.Windows.Forms.ToolStripMenuItem();
            this.cmbDestinationFolderPath = new System.Windows.Forms.ComboBox();
            this.wbExecutionResult = new System.Windows.Forms.WebBrowser();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.updateReportsBackgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.btnClear = new System.Windows.Forms.LinkLabel();
            this.btnCopyError = new System.Windows.Forms.LinkLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.btnOpenFile2File = new System.Windows.Forms.Button();
            this.btnOpenFile2Folder = new System.Windows.Forms.Button();
            this.cmbFile2Path = new System.Windows.Forms.ComboBox();
            this.btnSelectFile2 = new System.Windows.Forms.Button();
            this.lblFile2Path = new System.Windows.Forms.Label();
            this.btnOpenFile3File = new System.Windows.Forms.Button();
            this.btnOpenFile3Folder = new System.Windows.Forms.Button();
            this.cmbFile3Path = new System.Windows.Forms.ComboBox();
            this.btnSelectFile3 = new System.Windows.Forms.Button();
            this.lblFile3Path = new System.Windows.Forms.Label();
            this.btnOpenFile4File = new System.Windows.Forms.Button();
            this.btnOpenFile4Folder = new System.Windows.Forms.Button();
            this.cmbFile4Path = new System.Windows.Forms.ComboBox();
            this.btnSelectFile4 = new System.Windows.Forms.Button();
            this.lblFile4Path = new System.Windows.Forms.Label();
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
            this.statusStrip1.Location = new System.Drawing.Point(0, 781);
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
            // btnSelectFile1
            // 
            this.btnSelectFile1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFile1.Location = new System.Drawing.Point(982, 30);
            this.btnSelectFile1.Name = "btnSelectFile1";
            this.btnSelectFile1.Size = new System.Drawing.Size(29, 23);
            this.btnSelectFile1.TabIndex = 1;
            this.btnSelectFile1.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectFile1, "Apre la finestra di selezione file");
            this.btnSelectFile1.UseVisualStyleBackColor = true;
            this.btnSelectFile1.Click += new System.EventHandler(this.btnSelectReportFile_Click);
            // 
            // btnSelectDestinationFolder
            // 
            this.btnSelectDestinationFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectDestinationFolder.Location = new System.Drawing.Point(982, 149);
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
            this.lblCartellaOutputPath.Location = new System.Drawing.Point(12, 154);
            this.lblCartellaOutputPath.Name = "lblCartellaOutputPath";
            this.lblCartellaOutputPath.Size = new System.Drawing.Size(115, 13);
            this.lblCartellaOutputPath.TabIndex = 11;
            this.lblCartellaOutputPath.Text = "Cartella di destinazione";
            this.toolTipDefault.SetToolTip(this.lblCartellaOutputPath, "Cartella nella quale vettanno salvati i file di output");
            // 
            // btnOpenFile1File
            // 
            this.btnOpenFile1File.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFile1File.Image = global::GDPptGeneratorUI.Properties.Resources.png_clipart_spreadsheet_computer_icons_google_docs_microsoft_excel_table_angle_furniture_small1;
            this.btnOpenFile1File.Location = new System.Drawing.Point(1055, 30);
            this.btnOpenFile1File.Name = "btnOpenFile1File";
            this.btnOpenFile1File.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFile1File.TabIndex = 3;
            this.toolTipDefault.SetToolTip(this.btnOpenFile1File, "Apre il file report");
            this.btnOpenFile1File.UseVisualStyleBackColor = true;
            this.btnOpenFile1File.Click += new System.EventHandler(this.btnOpenReportFile_Click);
            // 
            // btnOpenDestFolder
            // 
            this.btnOpenDestFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenDestFolder.Image = global::GDPptGeneratorUI.Properties.Resources.free_yellow_open_folder_icon_11567_thumb_Small;
            this.btnOpenDestFolder.Location = new System.Drawing.Point(1019, 149);
            this.btnOpenDestFolder.Name = "btnOpenDestFolder";
            this.btnOpenDestFolder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenDestFolder.TabIndex = 6;
            this.toolTipDefault.SetToolTip(this.btnOpenDestFolder, "Apre la cartella di destinazione");
            this.btnOpenDestFolder.UseVisualStyleBackColor = true;
            this.btnOpenDestFolder.Click += new System.EventHandler(this.btnOpenDestFolder_Click);
            // 
            // btnOpenFile1Folder
            // 
            this.btnOpenFile1Folder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFile1Folder.Image = global::GDPptGeneratorUI.Properties.Resources.free_yellow_open_folder_icon_11567_thumb_Small;
            this.btnOpenFile1Folder.Location = new System.Drawing.Point(1019, 30);
            this.btnOpenFile1Folder.Name = "btnOpenFile1Folder";
            this.btnOpenFile1Folder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFile1Folder.TabIndex = 2;
            this.toolTipDefault.SetToolTip(this.btnOpenFile1Folder, "Apre la cartella contenente il file report");
            this.btnOpenFile1Folder.UseVisualStyleBackColor = true;
            this.btnOpenFile1Folder.Click += new System.EventHandler(this.btnOpenReportFolder_Click);
            // 
            // cmbFile1Path
            // 
            this.cmbFile1Path.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFile1Path.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbFile1Path.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbFile1Path.FormattingEnabled = true;
            this.cmbFile1Path.Location = new System.Drawing.Point(144, 32);
            this.cmbFile1Path.Name = "cmbFile1Path";
            this.cmbFile1Path.Size = new System.Drawing.Size(831, 21);
            this.cmbFile1Path.TabIndex = 0;
            this.cmbFile1Path.SelectedIndexChanged += new System.EventHandler(this.cmbReportFilePath_SelectedIndexChanged);
            this.cmbFile1Path.TextUpdate += new System.EventHandler(this.cmbReportFilePath_TextUpdate);
            // 
            // lblFile1Path
            // 
            this.lblFile1Path.AutoSize = true;
            this.lblFile1Path.Location = new System.Drawing.Point(12, 36);
            this.lblFile1Path.Name = "lblFile1Path";
            this.lblFile1Path.Size = new System.Drawing.Size(91, 13);
            this.lblFile1Path.TabIndex = 6;
            this.lblFile1Path.Text = "Percorso del file 1";
            // 
            // btnStart
            // 
            this.btnStart.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnStart.Location = new System.Drawing.Point(925, 190);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(153, 25);
            this.btnStart.TabIndex = 15;
            this.btnStart.Text = "Aggiorna report";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 202);
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
            this.toolStripMenuItemClear});
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(87, 20);
            this.toolStripMenuItem1.Text = "Impostazioni";
            // 
            // toolStripMenuItemClear
            // 
            this.toolStripMenuItemClear.Name = "toolStripMenuItemClear";
            this.toolStripMenuItemClear.Size = new System.Drawing.Size(256, 22);
            this.toolStripMenuItemClear.Text = "Svuota cronologia menu a tendina";
            this.toolStripMenuItemClear.Click += new System.EventHandler(this.toolStripMenuItemClear_Click);
            // 
            // cmbDestinationFolderPath
            // 
            this.cmbDestinationFolderPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbDestinationFolderPath.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbDestinationFolderPath.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbDestinationFolderPath.FormattingEnabled = true;
            this.cmbDestinationFolderPath.Location = new System.Drawing.Point(144, 150);
            this.cmbDestinationFolderPath.Name = "cmbDestinationFolderPath";
            this.cmbDestinationFolderPath.Size = new System.Drawing.Size(831, 21);
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
            this.wbExecutionResult.Location = new System.Drawing.Point(12, 227);
            this.wbExecutionResult.MinimumSize = new System.Drawing.Size(20, 20);
            this.wbExecutionResult.Name = "wbExecutionResult";
            this.wbExecutionResult.Size = new System.Drawing.Size(1066, 551);
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
            this.btnClear.Location = new System.Drawing.Point(938, 236);
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
            this.btnCopyError.AutoSize = true;
            this.btnCopyError.Location = new System.Drawing.Point(944, 647);
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
            // btnOpenFile2File
            // 
            this.btnOpenFile2File.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFile2File.Image = global::GDPptGeneratorUI.Properties.Resources.png_clipart_spreadsheet_computer_icons_google_docs_microsoft_excel_table_angle_furniture_small1;
            this.btnOpenFile2File.Location = new System.Drawing.Point(1055, 56);
            this.btnOpenFile2File.Name = "btnOpenFile2File";
            this.btnOpenFile2File.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFile2File.TabIndex = 25;
            this.toolTipDefault.SetToolTip(this.btnOpenFile2File, "Apre il file report");
            this.btnOpenFile2File.UseVisualStyleBackColor = true;
            // 
            // btnOpenFile2Folder
            // 
            this.btnOpenFile2Folder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFile2Folder.Image = global::GDPptGeneratorUI.Properties.Resources.free_yellow_open_folder_icon_11567_thumb_Small;
            this.btnOpenFile2Folder.Location = new System.Drawing.Point(1019, 56);
            this.btnOpenFile2Folder.Name = "btnOpenFile2Folder";
            this.btnOpenFile2Folder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFile2Folder.TabIndex = 24;
            this.toolTipDefault.SetToolTip(this.btnOpenFile2Folder, "Apre la cartella contenente il file report");
            this.btnOpenFile2Folder.UseVisualStyleBackColor = true;
            // 
            // cmbFile2Path
            // 
            this.cmbFile2Path.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFile2Path.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbFile2Path.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbFile2Path.FormattingEnabled = true;
            this.cmbFile2Path.Location = new System.Drawing.Point(144, 58);
            this.cmbFile2Path.Name = "cmbFile2Path";
            this.cmbFile2Path.Size = new System.Drawing.Size(831, 21);
            this.cmbFile2Path.TabIndex = 22;
            // 
            // btnSelectFile2
            // 
            this.btnSelectFile2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFile2.Location = new System.Drawing.Point(982, 56);
            this.btnSelectFile2.Name = "btnSelectFile2";
            this.btnSelectFile2.Size = new System.Drawing.Size(29, 23);
            this.btnSelectFile2.TabIndex = 23;
            this.btnSelectFile2.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectFile2, "Apre la finestra di selezione file");
            this.btnSelectFile2.UseVisualStyleBackColor = true;
            // 
            // lblFile2Path
            // 
            this.lblFile2Path.AutoSize = true;
            this.lblFile2Path.Location = new System.Drawing.Point(12, 62);
            this.lblFile2Path.Name = "lblFile2Path";
            this.lblFile2Path.Size = new System.Drawing.Size(91, 13);
            this.lblFile2Path.TabIndex = 26;
            this.lblFile2Path.Text = "Percorso del file 1";
            // 
            // btnOpenFile3File
            // 
            this.btnOpenFile3File.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFile3File.Image = global::GDPptGeneratorUI.Properties.Resources.png_clipart_spreadsheet_computer_icons_google_docs_microsoft_excel_table_angle_furniture_small1;
            this.btnOpenFile3File.Location = new System.Drawing.Point(1055, 82);
            this.btnOpenFile3File.Name = "btnOpenFile3File";
            this.btnOpenFile3File.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFile3File.TabIndex = 30;
            this.toolTipDefault.SetToolTip(this.btnOpenFile3File, "Apre il file report");
            this.btnOpenFile3File.UseVisualStyleBackColor = true;
            // 
            // btnOpenFile3Folder
            // 
            this.btnOpenFile3Folder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFile3Folder.Image = global::GDPptGeneratorUI.Properties.Resources.free_yellow_open_folder_icon_11567_thumb_Small;
            this.btnOpenFile3Folder.Location = new System.Drawing.Point(1019, 82);
            this.btnOpenFile3Folder.Name = "btnOpenFile3Folder";
            this.btnOpenFile3Folder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFile3Folder.TabIndex = 29;
            this.toolTipDefault.SetToolTip(this.btnOpenFile3Folder, "Apre la cartella contenente il file report");
            this.btnOpenFile3Folder.UseVisualStyleBackColor = true;
            // 
            // cmbFile3Path
            // 
            this.cmbFile3Path.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFile3Path.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbFile3Path.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbFile3Path.FormattingEnabled = true;
            this.cmbFile3Path.Location = new System.Drawing.Point(144, 84);
            this.cmbFile3Path.Name = "cmbFile3Path";
            this.cmbFile3Path.Size = new System.Drawing.Size(831, 21);
            this.cmbFile3Path.TabIndex = 27;
            // 
            // btnSelectFile3
            // 
            this.btnSelectFile3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFile3.Location = new System.Drawing.Point(982, 82);
            this.btnSelectFile3.Name = "btnSelectFile3";
            this.btnSelectFile3.Size = new System.Drawing.Size(29, 23);
            this.btnSelectFile3.TabIndex = 28;
            this.btnSelectFile3.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectFile3, "Apre la finestra di selezione file");
            this.btnSelectFile3.UseVisualStyleBackColor = true;
            // 
            // lblFile3Path
            // 
            this.lblFile3Path.AutoSize = true;
            this.lblFile3Path.Location = new System.Drawing.Point(12, 88);
            this.lblFile3Path.Name = "lblFile3Path";
            this.lblFile3Path.Size = new System.Drawing.Size(91, 13);
            this.lblFile3Path.TabIndex = 31;
            this.lblFile3Path.Text = "Percorso del file 1";
            // 
            // btnOpenFile4File
            // 
            this.btnOpenFile4File.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFile4File.Image = global::GDPptGeneratorUI.Properties.Resources.png_clipart_spreadsheet_computer_icons_google_docs_microsoft_excel_table_angle_furniture_small1;
            this.btnOpenFile4File.Location = new System.Drawing.Point(1055, 109);
            this.btnOpenFile4File.Name = "btnOpenFile4File";
            this.btnOpenFile4File.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFile4File.TabIndex = 35;
            this.toolTipDefault.SetToolTip(this.btnOpenFile4File, "Apre il file report");
            this.btnOpenFile4File.UseVisualStyleBackColor = true;
            // 
            // btnOpenFile4Folder
            // 
            this.btnOpenFile4Folder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFile4Folder.Image = global::GDPptGeneratorUI.Properties.Resources.free_yellow_open_folder_icon_11567_thumb_Small;
            this.btnOpenFile4Folder.Location = new System.Drawing.Point(1019, 109);
            this.btnOpenFile4Folder.Name = "btnOpenFile4Folder";
            this.btnOpenFile4Folder.Size = new System.Drawing.Size(29, 23);
            this.btnOpenFile4Folder.TabIndex = 34;
            this.toolTipDefault.SetToolTip(this.btnOpenFile4Folder, "Apre la cartella contenente il file report");
            this.btnOpenFile4Folder.UseVisualStyleBackColor = true;
            // 
            // cmbFile4Path
            // 
            this.cmbFile4Path.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFile4Path.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbFile4Path.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.HistoryList;
            this.cmbFile4Path.FormattingEnabled = true;
            this.cmbFile4Path.Location = new System.Drawing.Point(144, 111);
            this.cmbFile4Path.Name = "cmbFile4Path";
            this.cmbFile4Path.Size = new System.Drawing.Size(831, 21);
            this.cmbFile4Path.TabIndex = 32;
            // 
            // btnSelectFile4
            // 
            this.btnSelectFile4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFile4.Location = new System.Drawing.Point(982, 109);
            this.btnSelectFile4.Name = "btnSelectFile4";
            this.btnSelectFile4.Size = new System.Drawing.Size(29, 23);
            this.btnSelectFile4.TabIndex = 33;
            this.btnSelectFile4.Text = "...";
            this.toolTipDefault.SetToolTip(this.btnSelectFile4, "Apre la finestra di selezione file");
            this.btnSelectFile4.UseVisualStyleBackColor = true;
            // 
            // lblFile4Path
            // 
            this.lblFile4Path.AutoSize = true;
            this.lblFile4Path.Location = new System.Drawing.Point(12, 115);
            this.lblFile4Path.Name = "lblFile4Path";
            this.lblFile4Path.Size = new System.Drawing.Size(91, 13);
            this.lblFile4Path.TabIndex = 36;
            this.lblFile4Path.Text = "Percorso del file 1";
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
            this.ClientSize = new System.Drawing.Size(1093, 803);
            this.Controls.Add(this.btnOpenFile4File);
            this.Controls.Add(this.btnOpenFile4Folder);
            this.Controls.Add(this.cmbFile4Path);
            this.Controls.Add(this.btnSelectFile4);
            this.Controls.Add(this.lblFile4Path);
            this.Controls.Add(this.btnOpenFile3File);
            this.Controls.Add(this.btnOpenFile3Folder);
            this.Controls.Add(this.cmbFile3Path);
            this.Controls.Add(this.btnSelectFile3);
            this.Controls.Add(this.lblFile3Path);
            this.Controls.Add(this.btnOpenFile2File);
            this.Controls.Add(this.btnOpenFile2Folder);
            this.Controls.Add(this.cmbFile2Path);
            this.Controls.Add(this.btnSelectFile2);
            this.Controls.Add(this.lblFile2Path);
            this.Controls.Add(this.btnCopyError);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.wbExecutionResult);
            this.Controls.Add(this.btnOpenFile1File);
            this.Controls.Add(this.btnOpenDestFolder);
            this.Controls.Add(this.btnOpenFile1Folder);
            this.Controls.Add(this.cmbDestinationFolderPath);
            this.Controls.Add(this.btnSelectDestinationFolder);
            this.Controls.Add(this.lblCartellaOutputPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.cmbFile1Path);
            this.Controls.Add(this.btnSelectFile1);
            this.Controls.Add(this.lblFile1Path);
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
        private System.Windows.Forms.ComboBox cmbFile1Path;
        private System.Windows.Forms.Button btnSelectFile1;
        private System.Windows.Forms.Label lblFile1Path;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemClear;
        private System.Windows.Forms.ComboBox cmbDestinationFolderPath;
        private System.Windows.Forms.Button btnSelectDestinationFolder;
        private System.Windows.Forms.Label lblCartellaOutputPath;
        private System.Windows.Forms.Button btnOpenFile1Folder;
        private System.Windows.Forms.Button btnOpenDestFolder;
        private System.Windows.Forms.Button btnOpenFile1File;
        private System.Windows.Forms.WebBrowser wbExecutionResult;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.ComponentModel.BackgroundWorker updateReportsBackgroundWorker;
        private System.Windows.Forms.LinkLabel btnClear;
        private System.Windows.Forms.LinkLabel btnCopyError;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar;
        private WK.Libraries.BetterFolderBrowserNS.BetterFolderBrowser bfbDestFolder;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        private System.Windows.Forms.ToolStripStatusLabel lblVersion;
        private System.Windows.Forms.Button btnOpenFile2File;
        private System.Windows.Forms.Button btnOpenFile2Folder;
        private System.Windows.Forms.ComboBox cmbFile2Path;
        private System.Windows.Forms.Button btnSelectFile2;
        private System.Windows.Forms.Label lblFile2Path;
        private System.Windows.Forms.Button btnOpenFile3File;
        private System.Windows.Forms.Button btnOpenFile3Folder;
        private System.Windows.Forms.ComboBox cmbFile3Path;
        private System.Windows.Forms.Button btnSelectFile3;
        private System.Windows.Forms.Label lblFile3Path;
        private System.Windows.Forms.Button btnOpenFile4File;
        private System.Windows.Forms.Button btnOpenFile4Folder;
        private System.Windows.Forms.ComboBox cmbFile4Path;
        private System.Windows.Forms.Button btnSelectFile4;
        private System.Windows.Forms.Label lblFile4Path;
    }
}

