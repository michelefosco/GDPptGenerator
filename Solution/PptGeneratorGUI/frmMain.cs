using FilesEditor;
using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Windows.Forms;

namespace PptGeneratorGUI
{
    public partial class frmMain : Form
    {
        private DateTime _selectedDatePeriodo = DateTime.Today;
        private List<InputDataFilters_Item> _applicablefilters = new List<InputDataFilters_Item>();
        private bool _inputValidato = false;

        #region Files and folders paths
        private string SelectedFileBudgetPath
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
        private string SelectedFileForecastPath
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
        private string SelectedFileSuperDettagliPath
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
        private string SelectedFileRunRatePath
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
        private string SelectedDestinationFolderPath
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
        private string TmpFolderPath
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
                var debugFileName = $"DebugFile_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.xlsx";
                return Path.Combine(SelectedDestinationFolderPath, debugFileName);
            }
        }
        private string DataSourceFilePath
        {
            get
            {
                string exePath = Assembly.GetExecutingAssembly().Location;
                string exeDir = Path.GetDirectoryName(exePath);
                var sourceFilesFolder = Path.Combine(exeDir, "DataSourceFolder");
                return Path.Combine(sourceFilesFolder, FileNames.DATASOURCE_FILENAME);
            }
        }
        private string PowerPointTemplateFilePath
        {
            get
            {
                string exePath = Assembly.GetExecutingAssembly().Location;
                string exeDir = Path.GetDirectoryName(exePath);
                return Path.Combine(exeDir, FileNames.POWERPOINT_TEMPLATE_FILENAME);
            }
        }
        #endregion

        public frmMain()
        {
            InitializeComponent();

            FillComboBoxes();

            SetDefaultDatePeriodo();

            lblVersion.Text = $"Version: {GetVersion()}";
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
                _applicablefilters = new List<InputDataFilters_Item>();
            }

            gbPaths.Enabled = !_inputValidato;
            gbOptions.Enabled = _inputValidato;
            btnBuildPresentation.Enabled = _inputValidato;

            var isBudgetPathValid = IsBudgetPathValid();
            var isForecastPathValid = IsForecastPathValid();
            var isSuperDettagliPathValid = IsSuperDettagliPathValid();
            var isRunRatePathValid = IsRunRatePathValid();
            var isDestFolderValid = IsDestFolderValid();
            var allValid = isBudgetPathValid && isForecastPathValid && isSuperDettagliPathValid && isRunRatePathValid && isDestFolderValid;

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
            //
            btnValidaInput.Enabled = allValid && !_inputValidato;

            RefreshFiltersArea();

            if (_inputValidato)
            {
                toolTipDefault.SetToolTip(btnBuildPresentation, "Start generating the presentation");
                SetStatusLabel("Select input files and destination folder are validated. You can start building the presentation");
            }
            else if (allValid)
            {
                toolTipDefault.SetToolTip(btnValidaInput, "Start validation process");
                SetStatusLabel("Select input files and destination folder selected. You can start the validation process.");
            }
            else
            {
                toolTipDefault.SetToolTip(btnValidaInput, "Select input files and destination folder");
                SetStatusLabel("Select input files and destination folder");
            }
        }

        private void RefreshFiltersArea()
        {
            // posizione dei valori nelle colonne della griglia
            const int tableColumnIndex = 0;
            const int fieldColumnIndex = 1;
            const int selectButtonColumnIndex = 2;
            const int selectedValuesColumnIndex = 3;

            #region Aggiorno il contenuto della cella
            dgvFiltri.Enabled = false;
            dgvFiltri.Rows.Clear();

            if (_applicablefilters == null || _applicablefilters.Count == 0)
            {
                // esco lasciando la griglia vuota e disabilitata
                return;
            }

            foreach (var filtro in _applicablefilters.OrderBy(_ => _.Table.ToString()).ThenBy(_ => _.FieldName).ToList())
            {
                int rowIndex = dgvFiltri.Rows.Add();
                dgvFiltri.Rows[rowIndex].Cells[tableColumnIndex].Value = $"{filtro.Table}";
                dgvFiltri.Rows[rowIndex].Cells[fieldColumnIndex].Value = $"{filtro.FieldName}";
                dgvFiltri.Rows[rowIndex].Cells[selectButtonColumnIndex].Value = $"Select values";
                dgvFiltri.Rows[rowIndex].Cells[selectedValuesColumnIndex].Value = getTextForSelectedValueIntoTheFilter(filtro);
            }
            dgvFiltri.Enabled = true;
            #endregion

            #region Evento click sul pulsante
            dgvFiltri.CellContentClick += (s, e) =>
            {
                if (e.ColumnIndex == dgvFiltri.Columns["OpenFiltersSelection"].Index && e.RowIndex >= 0)
                {
                    // to prevent extra events during the selection of filters
                    dgvFiltri.Enabled = false;
                    var frmSelectFilters = new frmSelectFilters();

                    // assegno il filtro da modificare
                    var table = dgvFiltri.Rows[e.RowIndex].Cells[tableColumnIndex].Value.ToString();
                    var field = dgvFiltri.Rows[e.RowIndex].Cells[fieldColumnIndex].Value.ToString();
                    frmSelectFilters.FilterToManage = _applicablefilters.First(_ => _.Table.ToString() == table && _.FieldName == field);

                    // apro il form di selezione
                    frmSelectFilters.ShowDialog();

                    // refresh dei valori selezionati
                    dgvFiltri.Rows[e.RowIndex].Cells[selectedValuesColumnIndex].Value = getTextForSelectedValueIntoTheFilter(frmSelectFilters.FilterToManage);

                    // riabilito la griglia
                    dgvFiltri.Enabled = true;
                    frmSelectFilters.Dispose();
                }
            };
            #endregion

            // Organizza il contenuto della cella "Selected value" per il filtro
            string getTextForSelectedValueIntoTheFilter(InputDataFilters_Item filter)
            {
                return (filter.SelectedValues.Count == 0)
                        ? Values.ALLFILTERSAPPLIED
                        : string.Join("; ", filter.SelectedValues);
            }
        }


        #region Gestione history combo boxes
        private PathsHistory _pathFileHistory;
        private void FillComboBoxes()
        {
            LoadFileHistory();

            // Disabilito le combobox durante il caricamento per preventire eventi indesiderati
            cmbFileBudgetPath.Enabled = false;
            cmbFileForecastPath.Enabled = false;
            cmbFileSuperDettagliPath.Enabled = false;
            cmbFileRunRatePath.Enabled = false;
            cmbDestinationFolderPath.Enabled = false;

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

            // Riattivo le combobox
            cmbFileBudgetPath.Enabled = true;
            cmbFileForecastPath.Enabled = true;
            cmbFileSuperDettagliPath.Enabled = true;
            cmbFileRunRatePath.Enabled = true;
            cmbDestinationFolderPath.Enabled = true;
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

            bool configValue;

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
                MessageBox.Show("Folder does not exist or is invalid.", "Invalid folder", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show("Percorso del file non valido o file inesistente", "File non valido", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion


        #region Selezione data periodo
        private void SetDefaultDatePeriodo()
        {
            _selectedDatePeriodo = DateTime.Today;
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
            validaFileDiInput();
        }

        private void validaFileDiInput()
        {
            btnValidaInput.Enabled = false;
            gbPaths.Enabled = false;
            lblElaborazioneInCorso.Visible = true;
            ClearOutputArea();

            // input per la chiamata al backend
            var validateSourceFilesInput = new ValidateSourceFilesInput(
                    // proprietà classe base
                    destinationFolder: SelectedDestinationFolderPath,
                    tmpFolder: TmpFolderPath,
                    dataSourceFilePath: DataSourceFilePath,
                    debugFilePath: DebugFilePath,
                    //
                    fileBudgetPath: SelectedFileBudgetPath,
                    fileForecastPath: SelectedFileForecastPath,
                    fileSuperDettagliPath: SelectedFileSuperDettagliPath,
                    fileRunRatePath: SelectedFileRunRatePath);
            try
            {
                validaInputBackgroundWorker.RunWorkerAsync(validateSourceFilesInput);
            }
            catch (Exception ex)
            {
                showExpetion(ex);
                gbPaths.Enabled = true;
            }
        }

        private void validaInputBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var input = e.Argument as ValidateSourceFilesInput;
            var output = Editor.ValidateSourceFiles(input);
            e.Result = new object[] { input, output };
        }

        private void validaInputBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            var outputAndInput = e.Result as object[];
            var input = outputAndInput[0] as ValidateSourceFilesInput;
            var output = outputAndInput[1] as ValidateSourceFilesOutput;

            _inputValidato = output.Esito == EsitiFinali.Success;

            if (_inputValidato)
            {
                _applicablefilters = output.Applicablefilters;
                SetStatusLabel("Input validated successfully");
            }
            else
            {
                SetStatusLabel("Input validated with errors");
                SetOutputMessage(output.ManagedException);
                btnCopyError.Visible = true;
            }

            lblElaborazioneInCorso.Visible = false;
            RefreshUI(false);
        }
        #endregion


        #region CreaPresentazione
        private void btnBuildPresentation_Click(object sender, EventArgs e)
        {
            buildPresentation();
        }

        private void buildPresentation()
        {
            ClearOutputArea();

            btnBuildPresentation.Enabled = false;

            // Salvataggio dei percorsi selezionati
            AddPathsInXmlFileHistory();
            // Ricaricamento delle combobox
            FillComboBoxes();

            SetStatusLabel("Processing in progress...");
            lblElaborazioneInCorso.Visible = true;
            lblResults.Visible = false;
            Application.DoEvents();

            var buildPresentationInput = new BuildPresentationInput(
                // proprietà classe base
                dataSourceFilePath: DataSourceFilePath,
                destinationFolder: SelectedDestinationFolderPath,
                tmpFolder: TmpFolderPath,
                debugFilePath: DebugFilePath,
                //
                fileBudgetPath: SelectedFileBudgetPath,
                fileForecastPath: SelectedFileForecastPath,
                fileSuperDettagliPath: SelectedFileSuperDettagliPath,
                fileRunRatePath: SelectedFileRunRatePath,
                //
                powerPointTemplateFilePath: PowerPointTemplateFilePath,
                appendCurrentYear_FileSuperDettagli: cbAppendCurrentYearSuperDettagli.Checked,
                periodDate: _selectedDatePeriodo,
                applicablefilters: _applicablefilters
                );

            try
            {
                //buildPresentationBackgroundWorker.RunWorkerAsync(buildPresentationInput);
                Application.DoEvents();
                var output = Editor.BuildPresentation(buildPresentationInput);
                //  toolStripProgressBar.Visible = false;
                btnBuildPresentation.Enabled = true;
                lblElaborazioneInCorso.Visible = false;
                lblResults.Visible = true;
                Application.DoEvents();

                if (output.Esito == EsitiFinali.Success)
                {
                    string message = CreateOutputMessageSuccessHTML(output.DebugFilePath, output.OutputFilePathLists, output.Warnings);
                    SetOutputMessage(message);
                    SetStatusLabel("Processing completed successfully");
                }
                else //FAIL
                {
                    //Mostrare eventuali dati nel fail
                    SetStatusLabel("Processing completed with errors");
                    SetOutputMessage(output.ManagedException);
                    btnCopyError.Visible = true;
                }
            }
            catch (Exception ex)
            {
                btnBuildPresentation.Enabled = true;
                lblElaborazioneInCorso.Visible = false;
                lblResults.Visible = true;
                showExpetion(ex);
            }
        }

        private void buildPresentationBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            //var buildPresentationInput = e.Argument as BuildPresentationInput;
            //var output = Editor.BuildPresentation(buildPresentationInput);
            //e.Result = new object[] { buildPresentationInput, output };
        }

        private void buildPresentationBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //toolStripProgressBar.Visible = false;
            //btnBuildPresentation.Enabled = true;

            //// recupero input e output
            //var outputAndInput = e.Result as object[];
            //var input = outputAndInput[0] as BuildPresentationInput;
            //var output = outputAndInput[1] as BuildPresentationOutput;

            //if (output.Esito == EsitiFinali.Success)
            //{
            //    string message = CreateOutputMessageSuccessHTML(output.DebugFilePath, output.OutputFilePathLists, output.Warnings);
            //    SetOutputMessage(message);
            //    SetStatusLabel("Processing completed successfully");
            //    btnCopyError.Visible = false;
            //}
            //else //FAIL
            //{
            //    //Mostrare eventuali dati nel fail
            //    SetStatusLabel("Processing completed with errors");
            //    SetOutputMessage(output.ManagedException);
            //    btnCopyError.Visible = true;
            //}
        }
        #endregion




        #region Gestione output area
        private void showExpetion(Exception ex)
        {
            SetStatusLabel("Processing completed with errors.");

            if (ex is ManagedException mEx)
            { SetOutputMessage(mEx); }
            else
            { SetOutputMessage(ex); }

            btnCopyError.Visible = true;
        }

        private void SetOutputMessage(Exception ex)
        {
            if (ex == null) return;

            var htmlErrorMessage = HTML_Message_Helper.GetHTMLForExpetion(ex);
            SetOutputMessage(htmlErrorMessage);
        }

        private void SetOutputMessage(ManagedException mEx)
        {
            if (mEx == null) return;

            var htmlErrorMessage = HTML_Message_Helper.GetHTMLForExpetion(mEx);
            SetOutputMessage(htmlErrorMessage);
        }

        private void SetOutputMessage(string message)
        {
            var htmlMessage = HTML_Message_Helper.GetHTMLForBody(message);
            wbExecutionResult.DocumentText = null;
            Application.DoEvents();
            wbExecutionResult.DocumentText = htmlMessage;
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
            SetOutputMessage(HTML_Message_Helper._newlineHTML);
        }

        private string CreateOutputMessageSuccessHTML(string debugFilePath, List<string> outputFilePathLists, List<string> warnings)
        {
            string outputMessage = HTML_Message_Helper.GetHTMLGreenText(HTML_Message_Helper.GetHTMLBold("Processing completed successfully"));
            outputMessage += HTML_Message_Helper._newlineHTML;


            outputMessage += HTML_Message_Helper._newlineHTML;
            outputMessage += HTML_Message_Helper.GetHTMLBold("Presentations: ");
            // elenco delle presentazioni create
            foreach (var path in outputFilePathLists)
            {
                outputMessage += HTML_Message_Helper._newlineHTML;
                outputMessage += HTML_Message_Helper.GetHTMLHyperLink(path, path);
            }

            // link alla cartella di output
            var outputfolder = Path.GetDirectoryName(outputFilePathLists.First());
            outputMessage += HTML_Message_Helper._newlineHTML;
            outputMessage += HTML_Message_Helper._newlineHTML;
            outputMessage += HTML_Message_Helper.GetHTMLBold("Output folder: ");
            outputMessage += HTML_Message_Helper._newlineHTML;
            outputMessage += HTML_Message_Helper.GetHTMLHyperLink(outputfolder, outputfolder);


            // link al file di debug
            if (IsDebugModeEnabled())
            {
                outputMessage += HTML_Message_Helper._newlineHTML;
                outputMessage += HTML_Message_Helper._newlineHTML;
                outputMessage += HTML_Message_Helper.GetHTMLBold("Debug file: ");
                outputMessage += HTML_Message_Helper._newlineHTML;
                outputMessage += HTML_Message_Helper.GetHTMLHyperLink(debugFilePath, debugFilePath);
            }

            if (warnings.Count > 0)
            {
                outputMessage += HTML_Message_Helper._newlineHTML;
                outputMessage += HTML_Message_Helper._newlineHTML;
                outputMessage += HTML_Message_Helper.GeneraHtmlPerWarning(warnings);
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



        //private void wbExecutionResult_Navigating_1(object sender, WebBrowserNavigatingEventArgs e)
        //{
        //    string url = e.Url.OriginalString;
        //    if (url.StartsWith(HTML_Message_Helper.GetURLMarker(string.Empty)))
        //    {
        //        // Click su link file per l'apertura
        //        e.Cancel = true;

        //        url = HttpUtility.UrlDecode(url);
        //        url = url.Substring(HTML_Message_Helper.GetURLMarker(string.Empty).Length);


        //        if (File.Exists(url) || Directory.Exists(url))
        //        {
        //            System.Diagnostics.Process.Start(url);
        //        }
        //        else
        //        {
        //            //Se il file non è più esistente mostro un messaggio di errore
        //            MessageBox.Show($"Unable to open the file {url}; it is probably no longer present on the disk.", "Unable to open the file.", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        }

        //    }
        //    else if (url.StartsWith(HTML_Message_Helper.GetURLMarkerSetAsImput(string.Empty)))
        //    {
        //        e.Cancel = true;

        //        url = HttpUtility.UrlDecode(url);
        //        url = url.Substring(HTML_Message_Helper.GetURLMarkerSetAsImput(string.Empty).Length);

        //        RefreshUI(false);
        //    }
        //    //else if (url.StartsWith(GetURLMarkerDelete(string.Empty)))
        //    //{
        //    //    url = HttpUtility.UrlDecode(url);
        //    //    url = url.Substring(GetURLMarkerDelete(string.Empty).Length);

        //    //    e.Cancel = true;

        //    //    //Se il file non è più esistente mostro un messaggio di errore
        //    //    if (File.Exists(url))
        //    //        try
        //    //        {
        //    //            File.Delete(url);
        //    //            MessageBox.Show($"Il file {url} è stato eliminato.", "File eliminato", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    //        }
        //    //        catch 
        //    //        {
        //    //            MessageBox.Show($"Impossibile eliminare il file {url}, probabilmente è aperto, chiudere il file e riprovare.", "Impossibile eliminare il file", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    //        }
        //    //    else
        //    //        MessageBox.Show($"Impossibile eliminare il file {url}, probabilmente non è più presente sul disco.", "Impossibile eliminare il file", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    //}
        //}

        private void wbExecutionResult_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            string url = e.Url.OriginalString;
            if (url.StartsWith(HTML_Message_Helper.GetURLMarker(string.Empty)))
            {
                // Click su link file per l'apertura
                e.Cancel = true;

                url = HttpUtility.UrlDecode(url);
                url = url.Substring(HTML_Message_Helper.GetURLMarker(string.Empty).Length);


                if (File.Exists(url) || Directory.Exists(url))
                {
                    System.Diagnostics.Process.Start(url);
                }
                else
                {
                    //Se il file non è più esistente mostro un messaggio di errore
                    MessageBox.Show($"Unable to open the file {url}; it is probably no longer present on the disk.", "Unable to open the file.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            else if (url.StartsWith(HTML_Message_Helper.GetURLMarkerSetAsImput(string.Empty)))
            {
                e.Cancel = true;

                url = HttpUtility.UrlDecode(url);
                url = url.Substring(HTML_Message_Helper.GetURLMarkerSetAsImput(string.Empty).Length);

                RefreshUI(false);
            }
        }

        private void btnCopyError_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Clipboard.SetText(wbExecutionResult.Document.Body.InnerText);
            MessageBox.Show("Error copied to clipboard", "Error copied to clipboard", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        #endregion


        #region Click voci menu a tendina
        private void openDataSouceExcelFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openExcelForUser(DataSourceFilePath);
        }

        /// <summary>
        /// Click sul menù a tendina "Open source files folder"
        /// </summary>
        private void openSouceFilesFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var dataSourceFolderPath = Path.GetDirectoryName(DataSourceFilePath);
            openFolderForUser(dataSourceFolderPath);
        }

        private void loadLastSessionPathsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            LoadLastSessionFilePaths();
        }
        private void LoadLastSessionFilePaths()
        {
            if (cmbFileBudgetPath.Items.Count > 0)
            {   cmbFileBudgetPath.Text = cmbFileBudgetPath.Items[0].ToString(); }

            if (cmbFileForecastPath.Items.Count > 0)
            {   cmbFileForecastPath.Text = cmbFileForecastPath.Items[0].ToString(); }

            if (cmbFileRunRatePath.Items.Count > 0)
            {   cmbFileRunRatePath.Text = cmbFileRunRatePath.Items[0].ToString(); }

            if (cmbFileSuperDettagliPath.Items.Count > 0)
            {   cmbFileSuperDettagliPath.Text = cmbFileSuperDettagliPath.Items[0].ToString(); }

            if (cmbDestinationFolderPath.Items.Count > 0)
            {   cmbDestinationFolderPath.Text = cmbDestinationFolderPath.Items[0].ToString(); }
        }

        private void cleanCurrentsessionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cleanCurrentsession();
        }
        private void cleanCurrentsession()
        {
            _inputValidato = false;
            _applicablefilters = new List<InputDataFilters_Item>();
            _selectedDatePeriodo = DateTime.Today;

            ClearOutputArea();

            gbPaths.Enabled = true;

            SelectedFileBudgetPath = string.Empty;
            SelectedFileForecastPath = string.Empty;
            SelectedFileSuperDettagliPath = string.Empty;
            SelectedFileRunRatePath = string.Empty;
            SelectedDestinationFolderPath = string.Empty;

            cbAppendCurrentYearSuperDettagli.Checked = true;

            RefreshUI(true);
        }

        private void deletePathsHistoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DeletePathsHistory();
        }
                private void DeletePathsHistory()
        {
            if (MessageBox.Show("Permanently clear file history? (The operation is irreversible)", "Clear file history?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                _pathFileHistory.ClearHistory();
                FillComboBoxes();
                RefreshUI(false);
            }
        }
        #endregion
    }
}