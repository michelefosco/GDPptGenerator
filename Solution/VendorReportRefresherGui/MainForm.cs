using PptGenerator;
using PptGenerator.Entities;
using PptGenerator.Entities.Exceptions;
using PptGenerator.Enums;
using ReportRefresher.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Web;
using System.Windows.Forms;

namespace VendorReportRefresher
{
    public partial class MainForm : Form
    {
        private PathsHistory _pathFileHistory;
        private string _generatedReportFileName;
        private string _debugFileName;
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

        public string SelectedDestinationFilePath
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



        public string SelectedReportFilePath
        {
            get
            {
                return cmbFile1Path.Text;
            }
            set
            {
                cmbFile1Path.Text = value;
            }
        }

        public MainForm()
        {
            InitializeComponent();

            FillComboBoxes();

            lblVersion.Text = $"Versione: {GetVersion()}";
        }

        private void SetStatusLabel(string status)
        {
            txtStatusLabel.Text = status;
        }

        private void RefreshUI()
        {
            RefreshUI(false);
        }

        private void RefreshUI(bool breackRecursion)
        {
            bool isReportPathValid = IsReportPathValid();
            bool isDestFolderValid = IsDestFolderValid();

            if (isReportPathValid && isDestFolderValid)
            {
                toolTipDefault.SetToolTip(btnStart, "Avvia l'elaborazione del report");
                SetStatusLabel("File di input e cartella di destinazione selezionati, è possibile avviare l'elaborazione");
            }
            else
            {
                toolTipDefault.SetToolTip(btnStart, "Selezionare il file controller, il file report e la cartella di destinazione");
                SetStatusLabel("Selezionare i file di input e la cartella di destinazione");
            }

            btnOpenFile1Folder.Enabled = isReportPathValid;
            btnOpenFile1File.Enabled = isReportPathValid;
            btnOpenDestFolder.Enabled = isDestFolderValid;

            //Imposto il path di destinazione con lo stesso path del file report selezionato aggiungendo la cartella Output (che viene creata se non esiste)
            if (!breackRecursion && IsReportPathValid())
                SetOutputFolder();
        }

        private void SetOutputFolder()
        {
            if (string.IsNullOrEmpty(SelectedDestinationFilePath))
            {
                string outputFolderPath = Path.Combine(Path.GetDirectoryName(SelectedReportFilePath), "Output");

                if (!Directory.Exists(outputFolderPath))
                {
                    Directory.CreateDirectory(outputFolderPath);
                }

                SelectedDestinationFilePath = outputFolderPath;

                RefreshUI(true);
            }
        }

        public string GetLocaLApplicationDataPath()
        {
            string localApplicationPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GDReportRefresher");

            if (!Directory.Exists(localApplicationPath))
                Directory.CreateDirectory(localApplicationPath);

            return localApplicationPath;
        }

        public string GetFileHistoryFileName()
        {
            return Path.Combine(GetLocaLApplicationDataPath(), "GDReportRefresherFileHistory.xml");
        }

        private void LoadFileHistory()
        {
            _pathFileHistory = new PathsHistory(GetFileHistoryFileName());
        }

        private void FillComboBoxes()
        {
            LoadFileHistory();


            cmbFile1Path.Items.Clear();
            cmbDestinationFolderPath.Items.Clear();


            cmbFile1Path.Items.AddRange(_pathFileHistory.ReportPaths.ToArray());
            cmbDestinationFolderPath.Items.AddRange(_pathFileHistory.DestFolderPaths.ToArray());
        }

        private void AddPathsInXmlFileHistory()
        {
            _pathFileHistory.AddPathsHistory("...", SelectedReportFilePath, SelectedDestinationFilePath);
        }



        private bool IsReportPathValid()
        {
            bool isValid = !string.IsNullOrEmpty(SelectedReportFilePath);

            if (isValid)
            {
                isValid = File.Exists(SelectedReportFilePath);
            }

            return isValid;
        }

        private bool IsDestFolderValid()
        {
            bool isValid = !string.IsNullOrEmpty(SelectedDestinationFilePath);

            if (isValid)
            {
                isValid = Directory.Exists(SelectedDestinationFilePath);
            }

            return isValid;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            RefreshUI();
        }

        private void btnSelectControllerFile_Click(object sender, EventArgs e)
        {
            //Controller
            openFileDialog.Title = "Seleziona il file Controller";
            DialogResult controllerFileDialogResult = openFileDialog.ShowDialog();

            if (controllerFileDialogResult == DialogResult.OK)
            {
                //SelectedControllerFilePath = openFileDialog.FileName;
                RefreshUI();
            }
        }

        private void btnSelectReportFile_Click(object sender, EventArgs e)
        {
            //Report
            openFileDialog.Title = "Seleziona il file report";
            DialogResult reportFileDialogResult = openFileDialog.ShowDialog();

            if (reportFileDialogResult == DialogResult.OK)
            {
                SelectedReportFilePath = openFileDialog.FileName;
                RefreshUI();
            }
        }

        private void btnSelectDestinationFolder_Click(object sender, EventArgs e)
        {
            DialogResult folderBrowserResult = bfbDestFolder.ShowDialog();

            if (folderBrowserResult == DialogResult.OK)
            {
                SelectedDestinationFilePath = bfbDestFolder.SelectedPath;
                RefreshUI();
            }
        }

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
            htmlErrorMessage += StringToHTML(GetHTMLBold(mEx.MessaggioPerUtente));

            if (!string.IsNullOrEmpty(mEx.PercorsoFile))
            {
                htmlErrorMessage += _newlineHTML;
                htmlErrorMessage += _newlineHTML;
                htmlErrorMessage += StringToHTML("File: ") + GetHTMLHyperLink(mEx.PercorsoFile, mEx.PercorsoFile);
                //htmlErrorMessage += _spaceHTML;
                //htmlErrorMessage += GetHTMLDeleteFileHyperLink(mEx.PercorsoFile);
            }
            else
            {
                switch (mEx.TipologiaCartella)
                {
                    case TipologiaCartelle.FileDiTipo1:
                        htmlErrorMessage += _newlineHTML;
                        htmlErrorMessage += _newlineHTML;
                        htmlErrorMessage += StringToHTML("File: ") + GetHTMLHyperLink(SelectedReportFilePath, SelectedReportFilePath);
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
            string tableHTML = GetHTMLTableRowWithCells("Tipologia errore:", mEx.TipologiaErrore.GetEnumDescription());
            tableHTML += GetHTMLTableRowWithCells("Tipologia cartella:", mEx.TipologiaCartella.GetEnumDescription());

            if (!string.IsNullOrEmpty(mEx.WorksheetName))
            {
                tableHTML += GetHTMLTableRowWithCells("Nome foglio:", mEx.WorksheetName);
            }

            if (mEx.ColonnaCella.HasValue && mEx.RigaCella.HasValue)
            {
                tableHTML += GetHTMLTableRowWithCells("Cella:", $"{mEx.NomeColonnaCella}{mEx.RigaCella.ToString()}");
            }
            else
            {
                if (mEx.ColonnaCella.HasValue)
                {
                    tableHTML += GetHTMLTableRowWithCells("Colonna:", mEx.NomeColonnaCella);
                }

                if (mEx.RigaCella.HasValue)
                {
                    tableHTML += GetHTMLTableRowWithCells("Riga:", mEx.RigaCella.ToString());
                }
            }

            if (mEx.NomeDatoErrore != NomiDatoErrore.None)
            {
                tableHTML += GetHTMLTableRowWithCells("Errore sul dato:", mEx.NomeDatoErrore.GetEnumDescription());
            }

            if (!string.IsNullOrEmpty(mEx.Dato))
            {
                tableHTML += GetHTMLTableRowWithCells("Valore:", mEx.Dato);
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

        private void GenerateReport()
        {
            if (  IsReportPathValid() && IsDestFolderValid())
            {
                ClearOutputArea();
                AddPathsInXmlFileHistory();
                FillComboBoxes();

                _generatedReportFileName = string.Empty;
                _debugFileName = string.Empty;


                cmbFile1Path.SelectedIndex = 0;

                //Esecuzione Refresher
                SetStatusLabel("Elaborazione in corso...");

                string destPath = SelectedDestinationFilePath;
                string newReportFileName = Path.Combine(destPath, $"File_Gestione_Esterni.xlsx");

                string debugFileName = Path.Combine(destPath, "DebugFile.xlsx");

                //Togliere dopo i test? Se la chiave di configurazione non esiste viene sempre saltata l'eliminazione dei file
                if (IsOverwriteEnabled())
                {
                    if (File.Exists(newReportFileName))
                    {
                        DeleteFile(newReportFileName);
                    }

                    if (File.Exists(debugFileName))
                    {
                        DeleteFile(debugFileName);
                    }
                }
                else
                {
                    bool reportFileExist = File.Exists(newReportFileName);
                    bool debugFileExist = File.Exists(debugFileName);

                    //Richiesta di sovrascrivere se i file esistono già
                    if (reportFileExist || debugFileExist)
                    {
                        string message = reportFileExist && debugFileExist ? "I file " : "Il file ";

                        message += reportFileExist ? newReportFileName : "";
                        message += reportFileExist && debugFileExist ? " e " : "";
                        message += debugFileExist ? debugFileName : "";

                        message += reportFileExist && debugFileExist
                            ? " esistono già\r\nSovrascrivere i file? (L'operazione è irreversibile)"
                            : " esiste già\r\nSovrascrivere il file? (L'operazione è irreversibile)";

                        string caption = reportFileExist && debugFileExist ? "Sovrascrivere i file " : "Sovrascrivere il file ";
                        caption += reportFileExist ? "Report" : "";
                        caption += reportFileExist && debugFileExist ? " e " : "";
                        caption += debugFileExist ? "Debug" : "";

                        if (MessageBox.Show(message, caption, MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            if (File.Exists(newReportFileName))
                            {
                                DeleteFile(newReportFileName);
                            }

                            if (File.Exists(debugFileName))
                            {
                                DeleteFile(debugFileName);
                            }
                        }
                    }
                }

                var updateReportsInput =  new UpdateReportsInput("..", SelectedReportFilePath, newReportFileName, debugFileName);


                try
                {
                    toolStripProgressBar.Visible = true;
                    btnStart.Enabled = false;
                    updateReportsBackgroundWorker.RunWorkerAsync(updateReportsInput);
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
                MessageBox.Show("Selezionare i file controller, report e la cartella di destinazione", "File controller, report o cartella di destinazione non validi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            GenerateReport();
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

        private string CreateOutputMessageSuccessHTML(string message, string controllerFile, string reportFile, string newReportFile, string debugFile, List<RigaSpeseSkippata> righeSkippate)
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

            if (righeSkippate != null && righeSkippate.Count > 0)
            {
                outputMessage += _newlineHTML;
                outputMessage += _newlineHTML;
                outputMessage += _newlineHTML;
                outputMessage += GetHTMLBold("Attenzione, Alcune righe sono state scartate.");
                outputMessage += _newlineHTML;
                outputMessage += GetHTMLBold("Maggiori dettagli nel file di debug o di seguito.");
                outputMessage += _newlineHTML;
                outputMessage += _newlineHTML;
                outputMessage += GetHTMLMoreDetailLink($"Mostra maggiori dettagli ({righeSkippate.Count})");
                outputMessage += _newlineHTML;

                string moreDetails = string.Empty;
                foreach (RigaSpeseSkippata rigaSkippata in righeSkippate)
                {
                    moreDetails += $"Nome foglio: {rigaSkippata.Foglio}, Cella: {((ColumnIDS)rigaSkippata.Colonna).ToString()}{rigaSkippata.Riga}, Dato: \"{rigaSkippata.DatoNonValido}\"";
                    moreDetails += _newlineHTML;
                }

                outputMessage += GetInvisibleSPAN(moreDetails);
            }

            return outputMessage;
        }

        private void cmbControllerFilePath_TextUpdate(object sender, EventArgs e)
        {
            RefreshUI();
        }

        private void cmbReportFilePath_TextUpdate(object sender, EventArgs e)
        {
            RefreshUI();
        }

        private void cmbControllerFilePath_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshUI();
        }

        private void cmbReportFilePath_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshUI();
        }

        private void toolStripMenuItemClear_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Svuotare definitivamente lo storico dei file? (L'operaizone è irreversibile)", "Svuotare lo stodico dei file?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                _pathFileHistory.ClearHistory();
                FillComboBoxes();
                RefreshUI();
            }

        }

        private void btnOpenDestFolder_Click(object sender, EventArgs e)
        {
            if (IsDestFolderValid())
            {
                System.Diagnostics.Process.Start(SelectedDestinationFilePath);
            }
            else
            {
                RefreshUI();
                MessageBox.Show("Cartella di destinazione non esistente o nonvalida", "Cartella di destinazione non valida", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cmbDestinationFolderPath_TextUpdate(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(SelectedDestinationFilePath))
                RefreshUI(true);
            else
                RefreshUI();
        }

        private void cmbDestinationFolderPath_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshUI();
        }



        private void btnOpenReportFolder_Click(object sender, EventArgs e)
        {
            if (IsReportPathValid())
            {
                System.Diagnostics.Process.Start(Path.GetDirectoryName(SelectedReportFilePath));
            }
            else
            {
                RefreshUI();
                MessageBox.Show("Percorso del file report non valido o file inesistente", "File report non valido", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnOpenReportFile_Click(object sender, EventArgs e)
        {
            if (IsReportPathValid())
            {
                System.Diagnostics.Process.Start(SelectedReportFilePath);
            }
            else
            {
                RefreshUI();
                MessageBox.Show("Percorso del file report non valido o file inesistente", "File report non valido", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            else if(url.StartsWith(GetURLMarkerSetAsImput(string.Empty)))
            {
                url = HttpUtility.UrlDecode(url);
                url = url.Substring(GetURLMarkerSetAsImput(string.Empty).Length);

                e.Cancel = true;

                SelectedReportFilePath = url;
                RefreshUI();
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

        private void updateReportsBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            UpdateReportsInput updateReportsInput = e.Argument as UpdateReportsInput;

            UpdateReportsOutput output = Refresher.UpdateReports(updateReportsInput);
            e.Result = new object[] { updateReportsInput, output };
        }

        private void updateReportsBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            toolStripProgressBar.Visible = false;
            btnStart.Enabled = true;

            object[] outputAndInput = e.Result as object[];

            UpdateReportsInput updateReportsInput = outputAndInput[0] as UpdateReportsInput;
            UpdateReportsOutput output = outputAndInput[1] as UpdateReportsOutput;

 if (output.Esito == EsitiFinali.Success)
            {
                string message = CreateOutputMessageSuccessHTML("Elaborazione terminata con successo", "..", SelectedReportFilePath, updateReportsInput.NewReport_FilePath, updateReportsInput.FileDebug_FilePath, output.RigheSpesaSkippate);
                SetOutputMessage(message);
                SetStatusLabel("Elaborazione terminata con successo");

                _generatedReportFileName = updateReportsInput.NewReport_FilePath;
                _debugFileName = updateReportsInput.FileDebug_FilePath;
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

        private string GetVersion()
        {
            return Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }


        private void ResetSelectedReportAndDestFolder()
        {
            SelectedReportFilePath =
                SelectedDestinationFilePath = string.Empty;

            RefreshUI();
        }
    }
}
