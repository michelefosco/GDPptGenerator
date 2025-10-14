using DocumentFormat.OpenXml.Spreadsheet;
using FilesEditor;
using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using ShapeCrawler;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Remoting.Messaging;
using System.Web;
using System.Windows.Forms;

namespace PptGeneratorGUI
{
    public partial class MainForm : Form
    {

       // private string _debugFileName;
        private DateTime _selectedDatePeriodo;
        private List<InputDataFilters_Items> _applicablefilters;

        private bool _inputValidato = false;

        #region HTML elements
        private const string _newlineHTML = @"<BR />";
        private const string _boldHTML = @"<B>{0}</B>";
        private const string _hyperlinkHTML = @"<a style=""color: blue;"" href=""{0}"">{1}</a>";
        private const string _tabHTML = "&nbsp;&nbsp;&nbsp;";
        private const string _spaceHTML = "&nbsp;";
        private const string _redTextHTML = @"<span class=""red"">{0}</span>";
        private const string _greenTextHTML = @"<span class=""green"">{0}</span>";
        private const string _tableHTML = "<table>\r\n{0}\r\n</table>";
        private const string _trHTML = "  <tr>\r\n{0}\r\n  </tr>";
        private const string _tdHTML = "    <td>{0}</td>";
        private const string _invisibleSpanHTML = "<span id=\"invisibleSpan\" style=\"display: none;\">{0}</span>";
        private const string _moreDetailLink = @"<a href=""#"" style=""color: blue;"" onclick=""document.getElementById('invisibleSpan').style.display = 'inline'"">{0}</a>";
        private const string _deleteFileHyperlinkHTML = @"<a style=""color: red;"" href=""{0}"">{1}</a>";
        private const string _htmlBody =
 @"<html>
 <head>
  <style type=""text/css"">
  body {{
   font: 11px sans-serif;
  }}
  table {{
   font: 11px sans-serif;
  }}
th, td {{
  padding-right: 10px;
  }}
 .red {{
   font-family: sans-serif;
   color: red;
   font-size:13px;
  }}
  .green {{
   font-family: sans-serif;
   color: green;
   font-size:13px;
  }}
  </style>  
 </head>
 <body>
  {0}
 </body>
</html>";

        #endregion

        #region Selected paths
        public string SelectedFileBudgetPath
        {
            get
            {
                return cmbFileBudgetPath.Text;
            }
            set
            {
                cmbFileBudgetPath.Text = value;
            }
        }
        public string SelectedFileForecastPath
        {
            get
            {
                return cmbFileForecastPath.Text;
            }
            set
            {
                cmbFileForecastPath.Text = value;
            }
        }
        public string SelectedFileSuperDettagliPath
        {
            get
            {
                return cmbFileSuperDettagliPath.Text;
            }
            set
            {
                cmbFileSuperDettagliPath.Text = value;
            }
        }
        public string SelectedFileRunRatePath
        {
            get
            {
                return cmbFileRunRatePath.Text;
            }
            set
            {
                cmbFileRunRatePath.Text = value;
            }
        }
        public string SelectedDestinationFolderPath
        {
            get
            {
                return cmbDestinationFolderPath.Text;
            }
            set
            {
                cmbDestinationFolderPath.Text = value;
            }
        }

        private string SourceFilesFolderPath
        {
            get
            {
                string exePath = Assembly.GetExecutingAssembly().Location;
                string exeDir = Path.GetDirectoryName(exePath);
                var folderPath = Path.Combine(exeDir, "SourceFiles");
                return folderPath;
            }
        }


        private string TmpFolder
        {
            get
            {
                return Path.Combine(SelectedDestinationFolderPath, FolderNames.TMP_FOLDER_FOR_GENERATED_FILES); ;
            }
        }

        private string DebugFilePath
        {
            get
            {
                return Path.Combine(TmpFolder, "Debugfile.xlsx");
            }
        }


        #endregion

        public MainForm()
        {
            InitializeComponent();

            FillComboBoxes();

            SetDefaultsFor_ReplaceAll_CheckBoxes();

            SetDefaultDatePeriodo();

            lblVersion.Text = $"Versione: {GetVersion()}";
        }

        private void SetDefaultsFor_ReplaceAll_CheckBoxes()
        {
            cbReplaceAllDataFileSuperDettagli.Checked = true;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            RefreshUI(true);
        }


        private void SetStatusLabel(string status)
        {
            txtStatusLabel.Text = status;
        }

        private void RefreshUI(bool resetInputValidato)
        {
            if (resetInputValidato)
            {
                _inputValidato = false;
                gbOptions.Enabled = false;
                dgvFiltri.Rows.Clear();
            }

            bool isBudgetPathValid = IsBudgetPathValid();
            bool isForecastPathValid = IsForecastPathValid();
            bool isSuperDettagliPathValid = IsSuperDettagliPathValid();
            bool isRunRatePathValid = IsRunRatePathValid();
            bool isDestFolderValid = IsDestFolderValid();


            btnOpenFileBudgetFolder.Enabled = isBudgetPathValid;
            btnOpenFileBudget.Enabled = isBudgetPathValid;
            //
            btnOpenFileForecastFolder.Enabled = isForecastPathValid;
            btnOpenFileForecast.Enabled = isForecastPathValid;
            //
            btnOpenFileSuperDettagliFolder.Enabled = isSuperDettagliPathValid;
            btnOpenFileSuperDettagli.Enabled = isSuperDettagliPathValid;
            //
            btnOpenFileRunRateFolder.Enabled = isRunRatePathValid;
            btnOpenFileRunRate.Enabled = isRunRatePathValid;
            //
            btnOpenDestFolder.Enabled = isDestFolderValid;


            var allValid = isBudgetPathValid && isForecastPathValid && isSuperDettagliPathValid && isRunRatePathValid && isDestFolderValid;
            btnValidaInput.Enabled = allValid;

            if (allValid)
            {
                btnBuildPresentation.Enabled = _inputValidato;
                gbOptions.Enabled = _inputValidato;

                //todo:
                toolTipDefault.SetToolTip(btnBuildPresentation, "Avvia l'elaborazione del report");
                SetStatusLabel("File di input e cartella di destinazione selezionati, è possibile avviare l'elaborazione");
            }
            else
            {

                btnBuildPresentation.Enabled = false;
                gbOptions.Enabled = false;
                //todo:
                toolTipDefault.SetToolTip(btnBuildPresentation, "Selezionare il file controller, il file report e la cartella di destinazione");
                SetStatusLabel("Selezionare i file di input e la cartella di destinazione");
            }
        }

        private void BuildFiltersArea(List<InputDataFilters_Items> filtriPossibili)
        {
            dgvFiltri.Rows.Clear();

            foreach (var filtro in filtriPossibili)
            {
                int rowIndex = dgvFiltri.Rows.Add();
                dgvFiltri.Rows[rowIndex].Cells[0].Value = $"{filtro.Table}";

                dgvFiltri.Rows[rowIndex].Cells[1].Value = $"{filtro.FieldName}";

                dgvFiltri.Rows[rowIndex].Cells[2].Value = $"Select values";

                var textFiltriSelezionati = (filtro.SelectedValues.Count == 0)
                        ? FilesEditor.Constants.Values.ALLFILTERSAPPLIED
                        : string.Join("; ", filtro.SelectedValues);
                dgvFiltri.Rows[rowIndex].Cells[3].Value = textFiltriSelezionati;
            }

            dgvFiltri.CellContentClick += (s, e) =>
            {
                if (e.ColumnIndex == dgvFiltri.Columns["OpenFiltersSelection"].Index && e.RowIndex >= 0)
                {
                    //todo aprire form di selezione 
                    string tabella = dgvFiltri.Rows[e.RowIndex].Cells["Tabella"].Value.ToString();
                    string campo = dgvFiltri.Rows[e.RowIndex].Cells["Campo"].Value.ToString();
                    MessageBox.Show($"Hai cliccato sul pulsante della riga: {tabella}-{campo}");
                }
            };
        }


        #region Gestione history combo boxes
        private PathsHistory _pathFileHistory;
        private void FillComboBoxes()
        {
            LoadFileHistory();

            cmbFileBudgetPath.Items.Clear();
            cmbFileBudgetPath.Items.AddRange(_pathFileHistory.BudgetPaths.ToArray());

            cmbFileForecastPath.Items.Clear();
            cmbFileForecastPath.Items.AddRange(_pathFileHistory.ForecastPaths.ToArray());

            cmbFileSuperDettagliPath.Items.Clear();
            cmbFileSuperDettagliPath.Items.AddRange(_pathFileHistory.SuperDettagliPaths.ToArray());

            cmbFileRunRatePath.Items.Clear();
            cmbFileRunRatePath.Items.AddRange(_pathFileHistory.RunRatePaths.ToArray());

            cmbDestinationFolderPath.Items.Clear();
            cmbDestinationFolderPath.Items.AddRange(_pathFileHistory.DestFolderPaths.ToArray());
        }
        private void clearPathsHistoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Permanently clear file history? (The operation is irreversible)", "Clear file history?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                _pathFileHistory.ClearHistory();
                FillComboBoxes();
                RefreshUI(false);
            }
        }
        public string GetLocaLApplicationDataPath()
        {
            string localApplicationPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PptGenerator");

            if (!Directory.Exists(localApplicationPath))
            { Directory.CreateDirectory(localApplicationPath); }

            return localApplicationPath;
        }
        public string GetFileHistoryFileName()
        {
            return Path.Combine(GetLocaLApplicationDataPath(), "PptGeneratorFileHistory.xml");
        }
        private void LoadFileHistory()
        {
            _pathFileHistory = new PathsHistory(GetFileHistoryFileName());
        }

        private void AddPathsInXmlFileHistory()
        {
            _pathFileHistory.AddPathsHistory(SelectedFileBudgetPath, SelectedFileForecastPath, SelectedFileSuperDettagliPath, SelectedFileRunRatePath, SelectedDestinationFolderPath);
        }
        #endregion


        #region Check "IsValid" su file e cartella
        private bool IsBudgetPathValid()
        {
            bool isValid = !string.IsNullOrEmpty(SelectedFileBudgetPath);

            if (isValid)
            {
                isValid = File.Exists(SelectedFileBudgetPath);
            }

            return isValid;
        }

        private bool IsForecastPathValid()
        {
            bool isValid = !string.IsNullOrEmpty(SelectedFileForecastPath);

            if (isValid)
            {
                isValid = File.Exists(SelectedFileForecastPath);
            }

            return isValid;
        }

        private bool IsSuperDettagliPathValid()
        {
            bool isValid = !string.IsNullOrEmpty(SelectedFileSuperDettagliPath);

            if (isValid)
            {
                isValid = File.Exists(SelectedFileSuperDettagliPath);
            }

            return isValid;
        }

        private bool IsRunRatePathValid()
        {
            bool isValid = !string.IsNullOrEmpty(SelectedFileRunRatePath);

            if (isValid)
            {
                isValid = File.Exists(SelectedFileRunRatePath);
            }

            return isValid;
        }

        private bool IsDestFolderValid()
        {
            bool isValid = !string.IsNullOrEmpty(SelectedDestinationFolderPath);

            if (isValid)
            {
                isValid = Directory.Exists(SelectedDestinationFolderPath);
            }

            return isValid;
        }
        #endregion



        private bool IsDebugModeEnabled()
        {
            string debugModeEnabledValue = ConfigurationManager.AppSettings.Get("DebugModeEnabled");

            bool configValue = false;

            if (!string.IsNullOrEmpty(debugModeEnabledValue))
            {
                if (bool.TryParse(debugModeEnabledValue, out configValue))
                {
                    return configValue;
                }
                else
                    return false;
            }
            else
                return false;
        }



        private bool IsOverwriteEnabled()
        {
            string overwriteEnabledConfigValue = ConfigurationManager.AppSettings.Get("OverwriteOutputFile");

            bool configValue = false;

            if (!string.IsNullOrEmpty(overwriteEnabledConfigValue))
            {
                if (bool.TryParse(overwriteEnabledConfigValue, out configValue))
                {
                    return configValue;
                }
                else
                    return false;
            }
            else
                return false;
        }



        private string GetVersion()
        {
            return Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }



        #region Eventi Selezione file/cartella - RefreshUI
        private void cmbFileBudgetPath_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshUI(true);
        }
        private void cmbFileForecastPath_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshUI(true);
        }
        private void cmbFileSuperDettagliPath_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshUI(true);
        }
        private void cmbFileRunRatePath_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshUI(true);
        }
        private void cmbDestinationFolderPath_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshUI(true);
        }


        private void cmbFileBudgetPath_TextUpdate(object sender, EventArgs e)
        {
            RefreshUI(true);
        }
        private void cmbFileForecastPath_TextUpdate(object sender, EventArgs e)
        {
            RefreshUI(true);
        }
        private void cmbFileRunRatePath_TextUpdate(object sender, EventArgs e)
        {
            RefreshUI(true);
        }
        private void cmbFileSuperDettagliPath_TextUpdate(object sender, EventArgs e)
        {
            RefreshUI(true);
        }
        private void cmbDestinationFolderPath_TextUpdate(object sender, EventArgs e)
        {
            RefreshUI(true);
        }
        #endregion


        #region Eventi avvio selezione file/cartella
        private void btnSelectFileBudget_Click(object sender, EventArgs e)
        {
            const string tipoFoglio = "Budget";
            var title = $"Select the file {tipoFoglio}";
            var filePath = getPercosoSelezionatoDaUtente(title);
            if (!string.IsNullOrEmpty(filePath))
            {
                SelectedFileBudgetPath = filePath;
                RefreshUI(true);
            }
        }

        private void btnSelectForecastFile_Click(object sender, EventArgs e)
        {
            const string tipoFoglio = "Forecast";
            var title = $"Select the file {tipoFoglio}";
            var filePath = getPercosoSelezionatoDaUtente(title);
            if (!string.IsNullOrEmpty(filePath))
            {
                SelectedFileForecastPath = filePath;
                RefreshUI(true);
            }
        }

        private void btnSelectFileSuperDettagli_Click(object sender, EventArgs e)
        {
            const string tipoFoglio = "Super dettagli";
            var title = $"Select the file {tipoFoglio}";
            var filePath = getPercosoSelezionatoDaUtente(title);
            if (!string.IsNullOrEmpty(filePath))
            {
                SelectedFileSuperDettagliPath = filePath;
                RefreshUI(true);
            }
        }

        private void btnSelectFileRunRate_Click(object sender, EventArgs e)
        {
            const string tipoFoglio = "Run rate";
            var title = $"Select the file {tipoFoglio}";
            var filePath = getPercosoSelezionatoDaUtente(title);
            if (!string.IsNullOrEmpty(filePath))
            {
                SelectedFileRunRatePath = filePath;
                RefreshUI(true);
            }
        }

        private void btnSelectDestinationFolder_Click(object sender, EventArgs e)
        {
            DialogResult folderBrowserResult = bfbDestFolder.ShowDialog();

            if (folderBrowserResult == DialogResult.OK)
            {
                SelectedDestinationFolderPath = bfbDestFolder.SelectedPath;
                RefreshUI(false);
            }
        }

        private string getPercosoSelezionatoDaUtente(string title)
        {
            // Open the dialog window
            openFileDialog.Title = title;
            var fileDialogResult = openFileDialog.ShowDialog();

            // If the user selected a file, return the file path
            return (fileDialogResult == DialogResult.OK)
                ? openFileDialog.FileName
                : null;
        }
        #endregion


        #region Apertura cartelle selezionate
        private void btnOpenFileBudgetFolder_Click(object sender, EventArgs e)
        {
            var folderPath = Path.GetDirectoryName(SelectedFileBudgetPath);
            openFolderForUser(folderPath);
        }

        private void btnOpenFileForecastFolder_Click(object sender, EventArgs e)
        {
            var folderPath = Path.GetDirectoryName(SelectedFileForecastPath);
            openFolderForUser(folderPath);
        }

        private void btnOpenFileSuperDettagliFolder_Click(object sender, EventArgs e)
        {
            var folderPath = Path.GetDirectoryName(SelectedFileSuperDettagliPath);
            openFolderForUser(folderPath);
        }

        private void btnOpenFileRunRateFolder_Click(object sender, EventArgs e)
        {
            var folderPath = Path.GetDirectoryName(SelectedFileRunRatePath);
            openFolderForUser(folderPath);
        }

        private void btnOpenDestFolder_Click(object sender, EventArgs e)
        {
            openFolderForUser(SelectedDestinationFolderPath);
        }

        private void openFolderForUser(string folderPath)
        {
            if (Directory.Exists(folderPath))
            {
                System.Diagnostics.Process.Start(folderPath);
            }
            else
            {
                RefreshUI(false);
                MessageBox.Show("Cartella non esistente o non valida", "Cartella non valida", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion


        #region Apertura dei file Excel selezionati
        private void btnOpenFileBudget_Click(object sender, EventArgs e)
        {
            openExcelForUser(SelectedFileBudgetPath);
        }

        private void btnOpenFileForecast_Click(object sender, EventArgs e)
        {
            openExcelForUser(SelectedFileForecastPath);
        }

        private void btnOpenFileSuperDettagli_Click(object sender, EventArgs e)
        {
            openExcelForUser(SelectedFileSuperDettagliPath);
        }

        private void btnOpenFileRunRate_Click(object sender, EventArgs e)
        {
            openExcelForUser(SelectedFileRunRatePath);
        }

        private void openExcelForUser(string filePath)
        {
            if (File.Exists(filePath))
            {
                System.Diagnostics.Process.Start(filePath);
            }
            else
            {
                RefreshUI(false);
                MessageBox.Show("Percorso del file non valido o file inesistente", "File non valido", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion


        #region Selezione data periodo
        private void SetDefaultDatePeriodo()
        {
            _selectedDatePeriodo = DateTime.Now;
            lblDataPeriodo.Text = _selectedDatePeriodo.ToShortDateString();
        }

        private void btnOpenCalendar_Click(object sender, EventArgs e)
        {
            pnlCalendar.Visible = !pnlCalendar.Visible;
        }

        private void calendarPeriodo_DateSelected(object sender, DateRangeEventArgs e)
        {
            _selectedDatePeriodo = calendarPeriodo.SelectionStart;
            lblDataPeriodo.Text = calendarPeriodo.SelectionStart.ToShortDateString();
            pnlCalendar.Visible = false;
        }

        #endregion


        #region Valida Input
        private void btnValidaInput_Click(object sender, EventArgs e)
        {
            btnValidaInput.Enabled = false;
            validaFileDiInput();
        }

        private void validaFileDiInput()
        {
            // eseguzione dell'attività
            //btnNextBackgroundWorker.DoWork += (object sender, DoWorkEventArgs e) =>
            //{
            //    try
            //    {
            //        var input = e.Argument as ValidaSourceFilesInput;
            //        var output = Editor.ValidaSourceFiles(input);
            //        e.Result = new object[] { input, output };
            //    }
            //    //catch (ManagedException mEx)
            //    //{
            //    //    SetStatusLabel("Elaborazione terminata con errori");

            //    //    SetOutputMessage(mEx);
            //    //    btnCopyError.Visible = true;
            //    //}
            //    catch (Exception ex)
            //    {
            //        showExpetion(ex);
            //    }
            //};

            //// completamento dell'attività
            //btnNextBackgroundWorker.RunWorkerCompleted += (object sender, RunWorkerCompletedEventArgs e) =>
            //{
            //    try
            //    {
            //        var outputAndInput = e.Result as object[];
            //        var input = outputAndInput[0] as ValidaSourceFilesInput;
            //        var output = outputAndInput[1] as ValidaSourceFilesOutput;

            //        btnNext.Enabled = true;

            //        //todo valida input
            //        _inputValidato = true;

            //        if (_inputValidato)
            //        {
            //            _fieldFilters = output.UserOptions.Applicablefilters;
            //            BuildFiltersArea(_fieldFilters);
            //        }
            //        RefreshUI(false);
            //    }
            //    catch (ManagedException mEx)
            //    {
            //        SetStatusLabel("Elaborazione terminata con errori");

            //        SetOutputMessage(mEx);
            //        btnCopyError.Visible = true;
            //    }
            //    catch (Exception ex)
            //    {
            //        showExpetion(ex);

            //        SetStatusLabel("Elaborazione terminata con errori");

            //        SetOutputMessage(ex);
            //        btnCopyError.Visible = true;
            //    }
            //};

            //toolStripProgressBar.Visible = true;

            // input per la chiamata al backend
            var validaSourceFilesInput = new ValidaSourceFilesInput(
                    // proprietà classe base
                    destinationFolder: SelectedDestinationFolderPath,
                    tmpFolder: TmpFolder,
                    sourceFilesFolderPath: SourceFilesFolderPath,
                    fileDebugPath: DebugFilePath,
                    //
                    fileBudgetPath: SelectedFileBudgetPath,
                    fileForecastPath: SelectedFileForecastPath,
                    fileSuperDettagliPath: SelectedFileSuperDettagliPath,
                    fileRunRatePath: SelectedFileRunRatePath);
            try
            {
                validaInputBackgroundWorker.RunWorkerAsync(validaSourceFilesInput);
            }
            catch (Exception ex)
            {
                showExpetion(ex);
            }
        }

        private void validaInputBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var input = e.Argument as ValidaSourceFilesInput;
            var output = Editor.ValidaSourceFiles(input);
            e.Result = new object[] { input, output };
        }

        private void validaInputBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            var outputAndInput = e.Result as object[];
            var input = outputAndInput[0] as ValidaSourceFilesInput;
            var output = outputAndInput[1] as ValidaSourceFilesOutput;

            if (output.Esito == EsitiFinali.Success)
            {
                SetStatusLabel("Elaborazione terminata con successo");
                //todo valida input
                _inputValidato = true;
                _applicablefilters = output.Applicablefilters;
                BuildFiltersArea(_applicablefilters);
            }
            else
            {
                SetStatusLabel("Elaborazione terminata con errori");
                SetOutputMessage(output.ManagedException);
                btnCopyError.Visible = true;
                //todo valida input
                _inputValidato = false;
            }

            RefreshUI(resetInputValidato: false);
        }
        #endregion


        #region CreaPresentazione
        private void btnBuildPresentation_Click(object sender, EventArgs e)
        {
            buildPresentation();
        }

        private void buildPresentation()
        {
            //var backgroundWorker = new BackgroundWorker();

            //backgroundWorker.DoWork += (object sender, DoWorkEventArgs e) =>
            //{
            //    var buildPresentationInput = e.Argument as BuildPresentationInput;
            //    var output = Editor.BuildPresentation(buildPresentationInput);
            //    e.Result = new object[] { buildPresentationInput, output };
            //};

            //backgroundWorker.RunWorkerCompleted += (object sender, RunWorkerCompletedEventArgs e) =>
            //{
            //    toolStripProgressBar.Visible = false;
            //    btnCreaPresentazione.Enabled = true;

            //    var outputAndInput = e.Result as object[];

            //    var input = outputAndInput[0] as BuildPresentationInput;
            //    var output = outputAndInput[1] as BuildPresentationOutput;

            //    if (output.Esito == EsitiFinali.Success)
            //    {
            //        string message = CreateOutputMessageSuccessHTML("Elaborazione terminata con successo", "..", "SelectedReportFilePath", "updateReportsInput.NewReport_FilePath", input.FileDebug_FilePath/*, output.RigheSpesaSkippate*/);
            //        SetOutputMessage(message);
            //        SetStatusLabel("Elaborazione terminata con successo");

            //        // _generatedReportFileName = updateReportsInput.NewReport_FilePath;
            //        _debugFileName = input.FileDebug_FilePath;
            //        btnCopyError.Visible = false;
            //    }
            //    else //FAIL
            //    {
            //        //Mostrare eventuali dati nel fail
            //        SetStatusLabel("Elaborazione terminata con errori");
            //        SetOutputMessage(output.ManagedException);
            //        btnCopyError.Visible = true;
            //    }
            //};



            bool isBudgetPathValid = IsBudgetPathValid();
            bool isForecastPathValid = IsForecastPathValid();
            bool isSuperDettagliPathValid = IsSuperDettagliPathValid();
            bool isRunRatePathValid = IsRunRatePathValid();
            bool isDestFolderValid = IsDestFolderValid();

            if (/*isBudgetPathValid && isForecastPathValid && isSuperDettagliPathValid && isRunRatePathValid && */isDestFolderValid)
            {
                ClearOutputArea();
                AddPathsInXmlFileHistory();
                FillComboBoxes();


                cmbFileBudgetPath.SelectedIndex = 0;
                cmbFileForecastPath.SelectedIndex = 0;
                cmbFileSuperDettagliPath.SelectedIndex = 0;
                cmbFileRunRatePath.SelectedIndex = 0;
                cmbDestinationFolderPath.SelectedIndex = 0;

                //Esecuzione Refresher
                SetStatusLabel("Elaborazione in corso...");




                var buildPresentationInput = new BuildPresentationInput(
                    // proprietà classe base
                    destinationFolder: SelectedDestinationFolderPath,
                    tmpFolder: TmpFolder,
                    sourceFilesFolderPath: SourceFilesFolderPath,
                    fileDebugPath: DebugFilePath,
                    //
                    replaceAllData_FileSuperDettagli: cbReplaceAllDataFileSuperDettagli.Checked,
                    periodDate: _selectedDatePeriodo
                    );
                try
                {
                    toolStripProgressBar.Visible = true;
                    btnBuildPresentation.Enabled = false;
                    //btnBuildPresentationBackgroundWorker.RunWorkerAsync(buildPresentationInput);
                    buildPresentationBackgroundWorker.RunWorkerAsync(buildPresentationInput);
                }
                catch (ManagedException mEx)
                {
                    SetStatusLabel("Elaborazione terminata con errori");

                    SetOutputMessage(mEx);
                    btnCopyError.Visible = true;
                }
                catch (Exception ex)
                {
                    SetStatusLabel("Elaborazione terminata con errori");

                    SetOutputMessage(ex);
                    btnCopyError.Visible = true;
                }
            }
            else
            {
                MessageBox.Show("Selezionare i file di input e la cartella di destinazione", "File di input o cartella di destinazione non validi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buildPresentationBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var buildPresentationInput = e.Argument as BuildPresentationInput;
            var output = Editor.BuildPresentation(buildPresentationInput);
            e.Result = new object[] { buildPresentationInput, output };
        }

        private void buildPresentationBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            toolStripProgressBar.Visible = false;
            btnBuildPresentation.Enabled = true;

            var outputAndInput = e.Result as object[];

            var input = outputAndInput[0] as BuildPresentationInput;
            var output = outputAndInput[1] as BuildPresentationOutput;

            if (output.Esito == EsitiFinali.Success)
            {
                string message = CreateOutputMessageSuccessHTML("Elaborazione terminata con successo", "..", "SelectedReportFilePath", "updateReportsInput.NewReport_FilePath", input.FileDebugPath/*, output.RigheSpesaSkippate*/);
                SetOutputMessage(message);
                SetStatusLabel("Elaborazione terminata con successo");
                btnCopyError.Visible = false;
            }
            else //FAIL
            {
                //Mostrare eventuali dati nel fail
                SetStatusLabel("Elaborazione terminata con errori");
                SetOutputMessage(output.ManagedException);
                btnCopyError.Visible = true;
            }
        }


        //private void btnBuildPresentationBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        //{
        //    var buildPresentationInput = e.Argument as BuildPresentationInput;
        //    var output = Editor.BuildPresentation(buildPresentationInput);
        //    e.Result = new object[] { buildPresentationInput, output };
        //}

        //private void btnBuildPresentationBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        //{

        //    toolStripProgressBar.Visible = false;
        //    btnCreaPresentazione.Enabled = true;

        //    var outputAndInput = e.Result as object[];

        //    var input = outputAndInput[0] as BuildPresentationInput;
        //    var output = outputAndInput[1] as BuildPresentationOutput;

        //    if (output.Esito == EsitiFinali.Success)
        //    {
        //        string message = CreateOutputMessageSuccessHTML("Elaborazione terminata con successo", "..", "SelectedReportFilePath", "updateReportsInput.NewReport_FilePath", input.FileDebug_FilePath, output.RigheSpesaSkippate);
        //        SetOutputMessage(message);
        //        SetStatusLabel("Elaborazione terminata con successo");

        //        // _generatedReportFileName = updateReportsInput.NewReport_FilePath;
        //        _debugFileName = input.FileDebug_FilePath;
        //        btnCopyError.Visible = false;
        //    }
        //    else //FAIL
        //    {
        //        //Mostrare eventuali dati nel fail
        //        SetStatusLabel("Elaborazione terminata con errori");
        //        SetOutputMessage(output.ManagedException);
        //        btnCopyError.Visible = true;
        //    }
        //}
        #endregion


        private void openSouceFilesFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var folderPath = SourceFilesFolderPath;
            openFolderForUser(folderPath);
        }


        #region Gestione output area
        private void showExpetion(Exception ex)
        {
            // todo: translate
            SetStatusLabel("Processing completed with errors.");

            if (ex is ManagedException mEx)
            { SetOutputMessage(mEx); }
            else
            { SetOutputMessage(ex); }

            btnCopyError.Visible = true;
        }

        private void SetOutputMessage(string message)
        {
            string messageToHTML = message;
            string htmlMessage = string.Format(_htmlBody, messageToHTML);

            SetOutputMessageHTML(htmlMessage);
            btnClear.Visible = true;
        }

        private void SetOutputMessageHTML(string htmlMessage)
        {
            wbExecutionResult.DocumentText = htmlMessage;
        }

        private void SetOutputMessage(Exception ex)
        {
            string htmlErrorMessage = GetHTMLRedText(GetHTMLBold("Error:"));
            htmlErrorMessage += _newlineHTML;
            htmlErrorMessage += _newlineHTML;
            htmlErrorMessage += StringToHTML(ex.Message);
            htmlErrorMessage += _newlineHTML;
            htmlErrorMessage += _newlineHTML;

            htmlErrorMessage += GetInvisibleErrorDetails(ex);

            SetOutputMessage(htmlErrorMessage);
        }

        private void SetOutputMessage(ManagedException mEx)
        {
            string htmlErrorMessage = GetHTMLRedText(GetHTMLBold("Error:"));
            htmlErrorMessage += _newlineHTML;
            htmlErrorMessage += _newlineHTML;
            htmlErrorMessage += StringToHTML(GetHTMLBold(mEx.UserMessage));

            if (!string.IsNullOrEmpty(mEx.FilePath))
            {
                htmlErrorMessage += _newlineHTML;
                htmlErrorMessage += _newlineHTML;
                htmlErrorMessage += StringToHTML("File: ") + GetHTMLHyperLink(mEx.FilePath, mEx.FilePath);
                //htmlErrorMessage += _spaceHTML;
                //htmlErrorMessage += GetHTMLDeleteFileHyperLink(mEx.PercorsoFile);
            }
            else
            {
                switch (mEx.FileType)
                {
                    case FileTypes.DataSource_Template:
                        htmlErrorMessage += _newlineHTML;
                        htmlErrorMessage += _newlineHTML;
                        //htmlErrorMessage += StringToHTML("File: ") + GetHTMLHyperLink(SelectFileBudgetPath, "");
                        break;

                        //case TipologiaCartelle.FileDiTipo2:
                        //    htmlErrorMessage += _newlineHTML;
                        //    htmlErrorMessage += _newlineHTML;
                        //    htmlErrorMessage += StringToHTML("File: ") + GetHTMLHyperLink(SelectedControllerFilePath, SelectedControllerFilePath);
                        //    break;
                }
            }

            htmlErrorMessage += _newlineHTML;
            htmlErrorMessage += _newlineHTML;

            //Tabella con dati aggiuntivi dell'errore
            string tableHTML = GetHTMLTableRowWithCells("Error type: ", mEx.ErrorType.GetEnumDescription());
            tableHTML += GetHTMLTableRowWithCells("File type:", mEx.FileType.GetEnumDescription());

            if (!string.IsNullOrEmpty(mEx.WorksheetName))
            {
                tableHTML += GetHTMLTableRowWithCells("Worksheet name:", mEx.WorksheetName);
            }

            if (mEx.CellColumn.HasValue && mEx.CellRow.HasValue)
            {
                tableHTML += GetHTMLTableRowWithCells("Cell:", $"{((ColumnIDS)mEx.CellColumn).ToString()}{mEx.CellRow.ToString()}");
            }
            else
            {
                if (mEx.CellColumn.HasValue)
                {
                    tableHTML += GetHTMLTableRowWithCells("Column:", ((ColumnIDS)mEx.CellColumn).ToString());
                }

                if (mEx.CellRow.HasValue)
                {
                    tableHTML += GetHTMLTableRowWithCells("Row:", mEx.CellRow.ToString());
                }
            }

            //if (mEx.NomeDatoErrore != NomiDatoErrore.None)
            //{
            //    tableHTML += GetHTMLTableRowWithCells("Errore sul dato:", mEx.NomeDatoErrore.GetEnumDescription());
            //}

            if (!string.IsNullOrEmpty(mEx.Value))
            {
                tableHTML += GetHTMLTableRowWithCells("Value:", mEx.Value);
            }

            htmlErrorMessage += GetHTMLTable(tableHTML);

            htmlErrorMessage += _newlineHTML;
            htmlErrorMessage += _newlineHTML;
            htmlErrorMessage += GetInvisibleErrorDetails(mEx);

            SetOutputMessage(htmlErrorMessage);
        }

        private string StringToHTML(string str)
        {
            if (string.IsNullOrEmpty(str))
                return string.Empty;
            else
                return str.Replace("\t", _tabHTML)
                    .Replace(" ", _spaceHTML)
                    .Replace("\r\n", _newlineHTML)
                    .Replace("\n", _newlineHTML);
        }

        private string GetHTMLHyperLink(string url, string value)
        {
            return string.Format(_hyperlinkHTML, GetURLMarker(url), value);
        }

        private string GetHTMLHyperLinkSetAsImput(string url, string value)
        {
            return string.Format(_hyperlinkHTML, GetURLMarkerSetAsImput(url), value);
        }

        private string GetHTMLDeleteFileHyperLink(string url)
        {
            return string.Format(_deleteFileHyperlinkHTML, GetURLMarkerDelete(url), "(Clicca quì per cancellare il file)");
        }

        private string GetHTMLMoreDetailLink(string caption)
        {
            return string.Format(_moreDetailLink, caption);
        }

        private string GetHTMLBold(string str)
        {
            return string.Format(_boldHTML, str);
        }

        private string GetHTMLRedText(string str)
        {
            return string.Format(_redTextHTML, str);
        }

        private string GetHTMLGreenText(string str)
        {
            return string.Format(_greenTextHTML, str);
        }

        private string GetHTMLTable(string innerTableHTML)
        {
            return string.Format(_tableHTML, innerTableHTML);
        }

        private string GetHTMLTableRow(string innerRowHTML)
        {
            return string.Format(_trHTML, innerRowHTML);
        }

        private string GetHTMLTableCell(string innerCellHTML)
        {
            return string.Format(_tdHTML, innerCellHTML);
        }

        private string GetHTMLTableRowWithCells(string cell1Value, string cell2Value)
        {
            return GetHTMLTableRow(GetHTMLTableCell(cell1Value) + GetHTMLTableCell(GetHTMLBold(cell2Value)));
        }

        private void DeleteFile(string fullFileName)
        {
            try
            {
                File.Delete(fullFileName);
            }
            catch
            {
                MessageBox.Show($"Impossibile eliminare il file {fullFileName}, probabilmente è aperto o in uso, chiudere il file e riprovare.", "Impossibile eliminare il file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void ClearOutputArea()
        {
            if (wbExecutionResult.Document != null)
            {
                SetOutputMessage(" ");
            }

            btnClear.Visible = false;
            btnCopyError.Visible = false;
        }

        private string CreateOutputMessageSuccessHTMLDelete(string message, string nomeFornitore, string reportFile, string newReportFile)
        {
            string outputMessage = GetHTMLGreenText(GetHTMLBold(message));
            outputMessage += _newlineHTML;
            outputMessage += _newlineHTML;
            outputMessage += @"Eliminato il fornitore """;
            outputMessage += GetHTMLBold(nomeFornitore);
            outputMessage += @"""";
            outputMessage += _newlineHTML;
            outputMessage += GetHTMLBold("É stato generato il file: ");
            outputMessage += GetHTMLHyperLink(newReportFile, newReportFile);

            outputMessage += _newlineHTML;
            outputMessage += _newlineHTML;
            outputMessage += GetHTMLBold(GetHTMLHyperLinkSetAsImput(newReportFile, "Imposta il file generato come file di input"));

            return outputMessage;
        }

        private string CreateOutputMessageSuccessHTMLUpdate(string message, string nomeFornitore, string nuovoNomeFornitore, string reportFile, string newReportFile)
        {
            string outputMessage = GetHTMLGreenText(GetHTMLBold(message));
            outputMessage += _newlineHTML;
            outputMessage += _newlineHTML;
            outputMessage += @"Modificato il nome del fornitore """;
            outputMessage += GetHTMLBold(nomeFornitore);
            outputMessage += @""" in """;
            outputMessage += GetHTMLBold(nuovoNomeFornitore);
            outputMessage += @"""";
            outputMessage += _newlineHTML;
            outputMessage += GetHTMLBold("É stato generato il file: ");
            outputMessage += GetHTMLHyperLink(newReportFile, newReportFile);

            outputMessage += _newlineHTML;
            outputMessage += _newlineHTML;
            outputMessage += GetHTMLBold(GetHTMLHyperLinkSetAsImput(newReportFile, "Imposta il file generato come file di input"));

            return outputMessage;
        }

        private string CreateOutputMessageSuccessHTML(string message, string controllerFile, string reportFile, string newReportFile, string debugFile/*, List<RigaSpeseSkippata> righeSkippate*/)
        {
            string outputMessage = GetHTMLGreenText(GetHTMLBold(message));
            outputMessage += _newlineHTML;
            outputMessage += _newlineHTML;

            outputMessage += GetHTMLBold("É stato generato il file: ");
            outputMessage += GetHTMLHyperLink(newReportFile, newReportFile);

            if (IsDebugModeEnabled())
            {
                outputMessage += _newlineHTML;
                outputMessage += _newlineHTML;
                outputMessage += GetHTMLBold("É stato generato il file di DEBUG: ");
                outputMessage += GetHTMLHyperLink(debugFile, debugFile);
            }

            //if (righeSkippate != null && righeSkippate.Count > 0)
            //{
            //    outputMessage += _newlineHTML;
            //    outputMessage += _newlineHTML;
            //    outputMessage += _newlineHTML;
            //    outputMessage += GetHTMLBold("Attenzione, Alcune righe sono state scartate.");
            //    outputMessage += _newlineHTML;
            //    outputMessage += GetHTMLBold("Maggiori dettagli nel file di debug o di seguito.");
            //    outputMessage += _newlineHTML;
            //    outputMessage += _newlineHTML;
            //    outputMessage += GetHTMLMoreDetailLink($"Mostra maggiori dettagli ({righeSkippate.Count})");
            //    outputMessage += _newlineHTML;

            //    string moreDetails = string.Empty;
            //    foreach (RigaSpeseSkippata rigaSkippata in righeSkippate)
            //    {
            //        moreDetails += $"Nome foglio: {rigaSkippata.Foglio}, Cella: {((ColumnIDS)rigaSkippata.Colonna).ToString()}{rigaSkippata.Riga}, Dato: \"{rigaSkippata.DatoNonValido}\"";
            //        moreDetails += _newlineHTML;
            //    }

            //    outputMessage += GetInvisibleSPAN(moreDetails);
            //}

            return outputMessage;
        }

        private void wbExecutionResult_Navigating_1(object sender, WebBrowserNavigatingEventArgs e)
        {
            string url = e.Url.OriginalString;
            if (url.StartsWith(GetURLMarker(string.Empty)))
            {
                url = HttpUtility.UrlDecode(url);
                url = url.Substring(GetURLMarker(string.Empty).Length);

                e.Cancel = true;

                //Se il file non è più esistente mostro un messaggio di errore
                if (File.Exists(url))
                    System.Diagnostics.Process.Start(url);
                else
                    MessageBox.Show($"Impossibile aprire il file {url}, probabilmente non è più presente sul disco.", "Impossibile aprire il file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (url.StartsWith(GetURLMarkerSetAsImput(string.Empty)))
            {
                url = HttpUtility.UrlDecode(url);
                url = url.Substring(GetURLMarkerSetAsImput(string.Empty).Length);

                e.Cancel = true;

                //SelectedReportFilePath = url;
                RefreshUI(false);
            }
            //else if (url.StartsWith(GetURLMarkerDelete(string.Empty)))
            //{
            //    url = HttpUtility.UrlDecode(url);
            //    url = url.Substring(GetURLMarkerDelete(string.Empty).Length);

            //    e.Cancel = true;

            //    //Se il file non è più esistente mostro un messaggio di errore
            //    if (File.Exists(url))
            //        try
            //        {
            //            File.Delete(url);
            //            MessageBox.Show($"Il file {url} è stato eliminato.", "File eliminato", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //        }
            //        catch 
            //        {
            //            MessageBox.Show($"Impossibile eliminare il file {url}, probabilmente è aperto, chiudere il file e riprovare.", "Impossibile eliminare il file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        }
            //    else
            //        MessageBox.Show($"Impossibile eliminare il file {url}, probabilmente non è più presente sul disco.", "Impossibile eliminare il file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}

        }

        private string GetURLMarker(string url)
        {
            return "file-" + url;
        }

        private string GetURLMarkerSetAsImput(string url)
        {
            return "input-" + url;
        }

        private string GetURLMarkerDelete(string url)
        {
            return "filedelete-" + url;
        }

        private string GetInvisibleErrorDetails(Exception ex)
        {
            return GetInvisibleSPAN(StringToHTML($"Versione: {GetVersion()}\r\nErrore completo:\r\n{ex}"));
        }

        private string GetInvisibleSPAN(string innerHtml)
        {
            return string.Format(_invisibleSpanHTML, innerHtml);
        }

        private void btnCopyError_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Windows.Forms.Clipboard.SetText(wbExecutionResult.Document.Body.InnerText);
            MessageBox.Show("Errore copiato negli appunti", "Errore copiato negli appunti", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnClear_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ClearOutputArea();
        }


        private void ResetSelectedReportAndDestFolder()
        {
            SelectedFileBudgetPath = string.Empty;
            SelectedFileForecastPath = string.Empty;
            SelectedFileSuperDettagliPath = string.Empty;
            SelectedFileRunRatePath = string.Empty;
            SelectedDestinationFolderPath = string.Empty;

            RefreshUI(true);
        }

        #endregion
    }
}
