using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;

namespace FilesEditor.Steps.BuildPresentation
{
    internal class Step_LeggiInputFile_RanRate : StepBase
    {
        public Step_LeggiInputFile_RanRate(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            Context.DebugInfoLogger.LogStepContext("Step_LeggiInputFile_RanRate", Context);
            leggiInputFile();
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void leggiInputFile()
        {
            //
            var sourceFileType = FileTypes.RunRate;
            string sourceFilePath = Context.FileRunRatePath;
            string sourceWorksheetName = WorksheetNames.RUN_RATE_DATA;
            int souceHeaderRow = 1;
            int sourceHeaderFirstColumn = 1;
            //
            var destFileType = FileTypes.DataSource;
            string destFilePath = Context.DataSourceFilePath;
            string destWorksheetName = WorksheetNames.DATA_SOURCE_RUN_RATE_DATA;
            int destHeaderRow = 1;
            int destHeaderFirstColumn = 1;            
            //

            //var ePPlusHelperSource = GetHelperForExistingFile(pathFileSource, sourceFileType);
            //var ePPlusHelperDest = GetHelperForExistingFile(pathFileSource, destFileType);

            // todo: verifica che tutti gli headers nella destinazione siano presenti nella sorgente
            var packageSource = new ExcelPackage(new FileInfo(sourceFilePath));
            var packageDest = new ExcelPackage(new FileInfo(destFilePath));

            // Foglio sorgente e di destinazione
            var wsSource = packageSource.Workbook.Worksheets[sourceWorksheetName];
            var wsDest = packageDest.Workbook.Worksheets[destWorksheetName];

            // Lettura headers del foglio sorgente
            var sourceHeaders = new Dictionary<string, int>();
            int sourceCol = destHeaderFirstColumn;
            while (wsSource.Cells[souceHeaderRow, sourceCol].Value != null)
            {
                string header = wsSource.Cells[1, sourceCol].Text.Trim();
                sourceHeaders[header] = sourceCol;
                sourceCol++;
            }

            // Lettura headers del foglio di destinazione
            var destHeaders = new Dictionary<string, int>();
            int destCol = sourceHeaderFirstColumn;
            while (wsDest.Cells[destHeaderRow, destCol].Value != null)
            {
                string header = wsDest.Cells[1, destCol].Text.Trim();
                destHeaders[header] = destCol;
                destCol++;
            }

            var ePPlusHelperSource = GetHelperForExistingFile(sourceFilePath, sourceFileType);
            ThrowExpetionsForMissingHeader(ePPlusHelperSource, sourceWorksheetName, sourceFileType, souceHeaderRow, destHeaders.Keys.ToList());

            #region Verifico che il foglio sorgente abbia tutti gli headers necessari
            foreach (var kvp in destHeaders)
            {
                string destHeader = kvp.Key;

                // Se il foglio sorgente non ha tutti gli headers necessari sollevo l'eccezione
                if (!sourceHeaders.ContainsKey(destHeader))
                {
                    throw new ManagedException(
                        filePath:sourceFilePath,
                        fileType: sourceFileType,
                        //
                        worksheetName: sourceWorksheetName,
                        cellRow: souceHeaderRow,
                        cellColumn: null,
                        valueHeader: ValueHeaders.None,
                        value: destHeader,
                        //
                        errorType: ErrorTypes.MissingValue,
                        userMessage: string.Format(UserErrorMessages.MissingHeader, sourceFileType, destHeader, sourceWorksheetName)
                        );
                }
            }
            #endregion

            // Determina l'ultima riga con dati nel foglio sorgente
            int lastRowSource = wsSource.Dimension.End.Row;

            int destRowIndex = destHeaderRow + 1;

            // Copia dati riga per riga, rispettando i nomi delle colonne
            for (int rowSourceIndex = souceHeaderRow + 1; rowSourceIndex <= lastRowSource; rowSourceIndex++)
            {
                foreach (var kvp in destHeaders)
                {
                    string destHeader = kvp.Key;
                    int destColumnIndex = kvp.Value;

                    // Se il foglio sorgente ha la stessa colonna, copia il valore
                    if (sourceHeaders.ContainsKey(destHeader))
                    {
                        // trovo la colonna sorgente usando il nome della colonna di destinazione
                        int sourceColumnIndex = sourceHeaders[destHeader];

                        // prendo il valore dalla sorgete
                        var value = wsSource.Cells[rowSourceIndex, sourceColumnIndex].Value;

                        // lo scrivo nella destinazione
                        wsDest.Cells[destRowIndex, destColumnIndex].Value = value;
                    }
                }
            }

            // Salva le modifiche
            packageDest.Save();

            //// percorso cartella di backup
            //var sourceFilesFolder = Path.GetDirectoryName(Context.DataSourceFilePath);
            //var backupFolder = Path.Combine(sourceFilesFolder, Constants.FolderNames.DATASOURCE_FILES_BACKUP_FOLDER);
            //CreaDirectorySeNonEsiste(backupFolder); // creo la cartella            

            //// percorso file di backup
            //var backupFilePath = Path.Combine(backupFolder, $"{Path.GetFileNameWithoutExtension(Constants.FileNames.DATA_SOURCE_FILENAME)}_Backup_{System.DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(Constants.FileNames.DATA_SOURCE_FILENAME)}");

            //// copio il file
            //File.Copy(Context.DataSourceFilePath, backupFilePath, false);
        }
    }
}