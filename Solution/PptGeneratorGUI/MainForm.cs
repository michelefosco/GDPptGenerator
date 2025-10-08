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
using System.Web;
using System.Windows.Forms;

namespace PptGeneratorGUI
{
    public partial class MainForm : Form
    {
        private PathsHistory _pathFileHistory;
        private string _debugFileName;
        private DateTime _selectedDatePeriodo;
        private List<InputDataFilters_Items> _fieldFilters;

        private bool _inputValidato = false;

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
        public string SelectedFileRanRatePath
        {
            get
            {
                return cmbFileRanRatePath.Text;
            }
            set
            {
                cmbFileRanRatePath.Text = value;
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
        #endregion

        public MainForm()
        {
            InitializeComponent();

            FillComboBoxes();

            SetDefaultDatePeriodo();

            LetturaConfigurazioneDaFileDataSource();

            lblVersion.Text = $"Versione: {GetVersion()}";
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
            bool isRanRatePathValid = IsRanRatePathValid();
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
            btnOpenFileRanRateFolder.Enabled = isRanRatePathValid;
            btnOpenFileRanRate.Enabled = isRanRatePathValid;
            //
            btnOpenDestFolder.Enabled = isDestFolderValid;


            var allValid = isBudgetPathValid && isForecastPathValid && isSuperDettagliPathValid && isRanRatePathValid && isDestFolderValid;
            btnNext.Enabled = allValid;

            if (allValid)
            {
                btnCreaPresentazione.Enabled = _inputValidato;
                gbOptions.Enabled = _inputValidato;

                //todo:
                toolTipDefault.SetToolTip(btnCreaPresentazione, "Avvia l'elaborazione del report");
                SetStatusLabel("File di input e cartella di destinazione selezionati, è possibile avviare l'elaborazione");
            }
            else
            {

                btnCreaPresentazione.Enabled = false;
                gbOptions.Enabled = false;
                //todo:
                toolTipDefault.SetToolTip(btnCreaPresentazione, "Selezionare il file controller, il file report e la cartella di destinazione");
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
                        ? Values.ALLFILTERSAPPLIED
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

        private void LetturaConfigurazioneDaFileDataSource()
        {
            // throw new NotImplementedException();
        }


        #region Gestione history combo boxes
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

        private void FillComboBoxes()
        {
            LoadFileHistory();

            cmbFileBudgetPath.Items.Clear();
            cmbFileBudgetPath.Items.AddRange(_pathFileHistory.BudgetPaths.ToArray());

            cmbFileForecastPath.Items.Clear();
            cmbFileForecastPath.Items.AddRange(_pathFileHistory.ForecastPaths.ToArray());

            cmbFileSuperDettagliPath.Items.Clear();
            cmbFileSuperDettagliPath.Items.AddRange(_pathFileHistory.SuperDettagliPaths.ToArray());

            cmbFileRanRatePath.Items.Clear();
            cmbFileRanRatePath.Items.AddRange(_pathFileHistory.RanRatePaths.ToArray());

            cmbDestinationFolderPath.Items.Clear();
            cmbDestinationFolderPath.Items.AddRange(_pathFileHistory.DestFolderPaths.ToArray());
        }
        private void AddPathsInXmlFileHistory()
        {
            _pathFileHistory.AddPathsHistory(SelectedFileBudgetPath, SelectedFileForecastPath, SelectedFileSuperDettagliPath, SelectedFileRanRatePath, SelectedDestinationFolderPath);
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

        private bool IsRanRatePathValid()
        {
            bool isValid = !string.IsNullOrEmpty(SelectedFileRanRatePath);

            if (isValid)
            {
                isValid = File.Exists(SelectedFileRanRatePath);
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
            string htmlErrorMessage = GetHTMLRedText(GetHTMLBold("Errore:"));
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
            string htmlErrorMessage = GetHTMLRedText(GetHTMLBold("Errore:"));
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
            string tableHTML = GetHTMLTableRowWithCells("Tipologia errore:", mEx.ErrorType.GetEnumDescription());
            tableHTML += GetHTMLTableRowWithCells("Tipologia cartella:", mEx.FileType.GetEnumDescription());

            if (!string.IsNullOrEmpty(mEx.WorksheetName))
            {
                tableHTML += GetHTMLTableRowWithCells("Nome foglio:", mEx.WorksheetName);
            }

            if (mEx.CellColumn.HasValue && mEx.CellRow.HasValue)
            {
                tableHTML += GetHTMLTableRowWithCells("Cella:", $"{((ColumnIDS)mEx.CellColumn).ToString()}{mEx.CellRow.ToString()}");
            }
            else
            {
                if (mEx.CellColumn.HasValue)
                {
                    tableHTML += GetHTMLTableRowWithCells("Colonna:", ((ColumnIDS)mEx.CellColumn).ToString());
                }

                if (mEx.CellRow.HasValue)
                {
                    tableHTML += GetHTMLTableRowWithCells("Riga:", mEx.CellRow.ToString());
                }
            }

            //if (mEx.NomeDatoErrore != NomiDatoErrore.None)
            //{
            //    tableHTML += GetHTMLTableRowWithCells("Errore sul dato:", mEx.NomeDatoErrore.GetEnumDescription());
            //}

            if (!string.IsNullOrEmpty(mEx.Value))
            {
                tableHTML += GetHTMLTableRowWithCells("Valore:", mEx.Value);
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



        private void toolStripMenuItemClear_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Permanently clear file history? (The operation is irreversible)", "Clear file history?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                _pathFileHistory.ClearHistory();
                FillComboBoxes();
                RefreshUI(false);
            }
        }


        private void toolStripMenuItemOpenConfigFolder_Click(object sender, EventArgs e)
        {
            var folderPath = TemplatesFolderPath;
            openFolderForUser(folderPath);
        }



        private string TemplatesFolderPath
        {
            get
            {
                string exePath = Assembly.GetExecutingAssembly().Location;
                string exeDir = Path.GetDirectoryName(exePath);
                var folderPath = Path.Combine(exeDir, "Templates");
                return folderPath;
            }
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



        private string GetVersion()
        {
            return Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }


        private void ResetSelectedReportAndDestFolder()
        {
            SelectedFileBudgetPath = string.Empty;
            SelectedFileForecastPath = string.Empty;
            SelectedFileSuperDettagliPath = string.Empty;
            SelectedFileRanRatePath = string.Empty;
            SelectedDestinationFolderPath = string.Empty;

            RefreshUI(true);
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
        private void cmbFileRanRatePath_SelectedIndexChanged(object sender, EventArgs e)
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
        private void cmbFileRanRatePath_TextUpdate(object sender, EventArgs e)
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

        private void btnSelectFileRanRate_Click(object sender, EventArgs e)
        {
            const string tipoFoglio = "Ran rate";
            var title = $"Select the file {tipoFoglio}";
            var filePath = getPercosoSelezionatoDaUtente(title);
            if (!string.IsNullOrEmpty(filePath))
            {
                SelectedFileRanRatePath = filePath;
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

        private void btnOpenFileRanRateFolder_Click(object sender, EventArgs e)
        {
            var folderPath = Path.GetDirectoryName(SelectedFileRanRatePath);
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

        private void btnOpenFileRanRate_Click(object sender, EventArgs e)
        {
            openExcelForUser(SelectedFileRanRatePath);
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


        private void btnNext_Click(object sender, EventArgs e)
        {
            validaFileDiInput();
        }


        private void validaFileDiInput()
        {
            try
            {
                var backgroundWorker = new BackgroundWorker();

                // eseguzione dell'attività
                backgroundWorker.DoWork += (object sender, DoWorkEventArgs e) =>
                {
                    var input = e.Argument as ValidaSourceFilesInput;
                    var output = Editor.ValidaSourceFiles(input);
                    e.Result = new object[] { input, output };
                };

                // completamento dell'attività
                backgroundWorker.RunWorkerCompleted += (object sender, RunWorkerCompletedEventArgs e) =>
                {
                    var outputAndInput = e.Result as object[];
                    var input = outputAndInput[0] as ValidaSourceFilesInput;
                    var output = outputAndInput[1] as ValidaSourceFilesOutput;

                    btnNext.Enabled = true;

                    //todo valida input
                    _inputValidato = true;

                    if (_inputValidato)
                    {
                        _fieldFilters = output.UserOptions.Applicablefilters;
                        BuildFiltersArea(_fieldFilters);
                    }
                    RefreshUI(false);
                };

                toolStripProgressBar.Visible = true;
                btnNext.Enabled = false;
                // input per la chiamata al backend
                var validaSourceFilesInput = new ValidaSourceFilesInput(
                        destinationFolder: SelectedDestinationFolderPath,
                        templatesFolder: TemplatesFolderPath,
                        fileBudgetPath: SelectedFileBudgetPath,
                        fileForecastPath: SelectedFileForecastPath,
                        fileSuperDettagliPath: SelectedFileSuperDettagliPath,
                        fileRanRatePath: SelectedFileRanRatePath);
                //btnNextBackgroundWorker.RunWorkerAsync(getUserOptionsFromDataSourceInput);
                backgroundWorker.RunWorkerAsync(validaSourceFilesInput);
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

        //private void btnNextBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        //{
        //    var input = e.Argument as GetUserOptionsFromDataSourceInput;
        //    var output = Info.GetUserOptionsFromDataSource(input);

        //    e.Result = new object[] { input, output };
        //}

        //private void btnNextBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        //{
        //    btnNext.Enabled = true;

        //    //todo valida input
        //    _inputValidato = true;

        //    if (_inputValidato)
        //    {
        //        BuildFiltersArea(null);
        //        //

        //    }

        //    RefreshUI(false);
        //}


        #region CreaPresentazione
        private void btnCreaPresentazione_Click(object sender, EventArgs e)
        {
            createPresentation();
        }

        private void createPresentation()
        {
            var backgroundWorker = new BackgroundWorker();

            backgroundWorker.DoWork += (object sender, DoWorkEventArgs e) =>
            {
                var createPresentationsInput = e.Argument as CreatePresentationsInput;
                var output = Editor.CreatePresentations(createPresentationsInput);
                e.Result = new object[] { createPresentationsInput, output };
            };

            backgroundWorker.RunWorkerCompleted += (object sender, RunWorkerCompletedEventArgs e) =>
            {
                toolStripProgressBar.Visible = false;
                btnCreaPresentazione.Enabled = true;

                var outputAndInput = e.Result as object[];

                var input = outputAndInput[0] as CreatePresentationsInput;
                var output = outputAndInput[1] as CreatePresentationsOutput;

                if (output.Esito == EsitiFinali.Success)
                {
                    string message = CreateOutputMessageSuccessHTML("Elaborazione terminata con successo", "..", "SelectedReportFilePath", "updateReportsInput.NewReport_FilePath", input.FileDebug_FilePath/*, output.RigheSpesaSkippate*/);
                    SetOutputMessage(message);
                    SetStatusLabel("Elaborazione terminata con successo");

                    // _generatedReportFileName = updateReportsInput.NewReport_FilePath;
                    _debugFileName = input.FileDebug_FilePath;
                    btnCopyError.Visible = false;
                }
                else //FAIL
                {
                    //Mostrare eventuali dati nel fail
                    SetStatusLabel("Elaborazione terminata con errori");
                    SetOutputMessage(output.ManagedException);
                    btnCopyError.Visible = true;
                }
            };



            bool isBudgetPathValid = IsBudgetPathValid();
            bool isForecastPathValid = IsForecastPathValid();
            bool isSuperDettagliPathValid = IsSuperDettagliPathValid();
            bool isRanRatePathValid = IsRanRatePathValid();
            bool isDestFolderValid = IsDestFolderValid();

            if (/*isBudgetPathValid && isForecastPathValid && isSuperDettagliPathValid && isRanRatePathValid && */isDestFolderValid)
            {
                ClearOutputArea();
                AddPathsInXmlFileHistory();
                FillComboBoxes();


                cmbFileBudgetPath.SelectedIndex = 0;
                cmbFileForecastPath.SelectedIndex = 0;
                cmbFileSuperDettagliPath.SelectedIndex = 0;
                cmbFileRanRatePath.SelectedIndex = 0;
                cmbDestinationFolderPath.SelectedIndex = 0;

                //Esecuzione Refresher
                SetStatusLabel("Elaborazione in corso...");


                var tmpFolder = Path.Combine(SelectedDestinationFolderPath, FilesEditor.Constants.FolderNames.TMP_FOLDER_FOR_GENERATED_FILES);
                _debugFileName = Path.Combine(tmpFolder, "Debugfile.xlsx");

                var createPresentationsInput = new CreatePresentationsInput(
                            outputFolder: SelectedDestinationFolderPath,
                            tmpFolder: tmpFolder,
                            templatesFolder: TemplatesFolderPath,
                            fileDebug_FilePath: _debugFileName);
                try
                {
                    toolStripProgressBar.Visible = true;
                    btnCreaPresentazione.Enabled = false;
                    //btnCreatePresentationBackgroundWorker.RunWorkerAsync(createPresentationsInput);
                    backgroundWorker.RunWorkerAsync(createPresentationsInput);
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


        //private void btnCreatePresentationBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        //{
        //    var createPresentationsInput = e.Argument as CreatePresentationsInput;
        //    var output = Editor.CreatePresentations(createPresentationsInput);
        //    e.Result = new object[] { createPresentationsInput, output };
        //}

        //private void btnCreatePresentationBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        //{

        //    toolStripProgressBar.Visible = false;
        //    btnCreaPresentazione.Enabled = true;

        //    var outputAndInput = e.Result as object[];

        //    var input = outputAndInput[0] as CreatePresentationsInput;
        //    var output = outputAndInput[1] as CreatePresentationsOutput;

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
    }
}
