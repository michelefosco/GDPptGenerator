using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FilesEditor.Steps.BuildPresentation
{
    internal class Step_LeggiInputFiles : StepBase
    {
        public Step_LeggiInputFiles(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            Context.DebugInfoLogger.LogStepContext("Step_LeggiInputFiles", Context);

            leggiInputFile(
                    sourceFileType: FileTypes.Budget,
                    sourceFilePath: Context.FileBudgetPath,
                    sourceWorksheetName: WorksheetNames.BUDGET_DATA,
                    souceHeaderRow: Context.Configurazione.INPUT_FILES_BUDGET_HEADERS_ROW,
                    sourceHeaderFirstColumn: Context.Configurazione.INPUT_FILES_BUDGET_HEADERS_FIRST_COL,
                    //
                    destWorksheetName: WorksheetNames.DATA_SOURCE_BUDGET_DATA,
                    destHeaderRow: Context.Configurazione.DATASOURCE_BUDGET_HEADERS_ROW,
                    destHeaderFirstColumn: Context.Configurazione.DATASOURCE_BUDGET_HEADERS_FIRST_COL
                );

            leggiInputFile(
                    sourceFileType: FileTypes.Forecast,
                    sourceFilePath: Context.FileForecastPath,
                    sourceWorksheetName: WorksheetNames.FORECAST_DATA,
                    souceHeaderRow: Context.Configurazione.INPUT_FILES_FORECAST_HEADERS_ROW,
                    sourceHeaderFirstColumn: Context.Configurazione.INPUT_FILES_FORECAST_HEADERS_FIRST_COL,
                    //
                    destWorksheetName: WorksheetNames.DATA_SOURCE_FORECAST_DATA,
                    destHeaderRow: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_ROW,
                    destHeaderFirstColumn: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_FIRST_COL
                );

            leggiInputFile(
                    sourceFileType: FileTypes.RunRate,
                    sourceFilePath: Context.FileRunRatePath,
                    sourceWorksheetName: WorksheetNames.RUN_RATE_DATA,
                    souceHeaderRow: Context.Configurazione.INPUT_FILES_RUNRATE_HEADERS_ROW,
                    sourceHeaderFirstColumn: Context.Configurazione.INPUT_FILES_RUNRATE_HEADERS_FIRST_COL,
                    //
                    destWorksheetName: WorksheetNames.DATA_SOURCE_RUN_RATE_DATA,
                    destHeaderRow: Context.Configurazione.INPUT_FILES_RUNRATE_HEADERS_ROW,
                    destHeaderFirstColumn: Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_FIRST_COL
                );

            //todo: gestione particolare per Superdettagli
            leggiInputFile(
                    sourceFileType: FileTypes.SuperDettagli,
                    sourceFilePath: Context.FileSuperDettagliPath,
                    sourceWorksheetName: WorksheetNames.SUPERDETTAGLI_DATA,
                    souceHeaderRow: Context.Configurazione.INPUT_FILES_SUPERDETTAGLI_HEADERS_ROW,
                    sourceHeaderFirstColumn: Context.Configurazione.INPUT_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL,
                    //
                    destWorksheetName: WorksheetNames.DATA_SOURCE_SUPERDETTAGLI_DATA,
                    destHeaderRow: Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_ROW,
                    destHeaderFirstColumn: Context.Configurazione.INPUT_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL
                );

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void leggiInputFile(
                FileTypes sourceFileType,
                string sourceFilePath,
                string sourceWorksheetName,
                int souceHeaderRow,
                int sourceHeaderFirstColumn,
                //
                string destWorksheetName,
                int destHeaderRow,
                int destHeaderFirstColumn
            )
        {
            var packageSource = new ExcelPackage(new FileInfo(sourceFilePath));
            var packageDest = new ExcelPackage(new FileInfo(Context.DataSourceFilePath)); // Sempre datasource

            // Foglio sorgente e di destinazione
            var wsSource = packageSource.Workbook.Worksheets[sourceWorksheetName];
            var wsDest = packageDest.Workbook.Worksheets[destWorksheetName];

            //todo: refacotrying, estrapolare come metodo sull'helper
            // Lettura headers del foglio sorgente
            var sourceHeaders = new Dictionary<string, int>();
            var sourceCol = sourceHeaderFirstColumn;
            while (wsSource.Cells[souceHeaderRow, sourceCol].Value != null)
            {
                var header = wsSource.Cells[souceHeaderRow, sourceCol].Text.Trim().ToLower();
                sourceHeaders[header] = sourceCol;
                sourceCol++;
            }

            // Lettura headers del foglio di destinazione
            var destHeaders = new Dictionary<string, int>();
            var destCol = destHeaderFirstColumn;
            while (wsDest.Cells[destHeaderRow, destCol].Value != null)
            {
                string header = wsDest.Cells[destHeaderRow, destCol].Text.Trim().ToLower();
                destHeaders[header] = destCol;
                destCol++;
            }



            #region Verifico che tutti gli headers presenti nella destinazione sia effettivamente presenti nella sorgente
            foreach (var kvp in destHeaders)
            {
                string destHeader = kvp.Key;

                // Se il foglio sorgente non ha tutti gli headers necessari sollevo l'eccezione
                if (!sourceHeaders.ContainsKey(destHeader))
                {
                    throw new ManagedException(
                        filePath: sourceFilePath,
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



            int lastRowDest = wsDest.Dimension.End.Row;
            int lastColDest = wsDest.Dimension.End.Column;
            for (int rowDestIndex = destHeaderRow + 1; rowDestIndex <= lastRowDest; rowDestIndex++)
            {
                for (int colIndex = destHeaderFirstColumn; colIndex <= lastColDest; colIndex++)
                {
                    //ripulisco le righe esistenti
                    wsDest.Cells[rowDestIndex, colIndex].Clear();
                }
            }


            #region Preparo una struttura più snella che contenta le informazioni su filtri
            InputDataFilters_Tables inputDataFilters_Table;
            switch (sourceFileType)
            {
                case FileTypes.Budget:
                    inputDataFilters_Table = InputDataFilters_Tables.BUDGET;
                    break;
                case FileTypes.Forecast:
                    inputDataFilters_Table = InputDataFilters_Tables.FORECAST;
                    break;
                case FileTypes.RunRate:
                    inputDataFilters_Table = InputDataFilters_Tables.RUNRATE;
                    break;
                case FileTypes.SuperDettagli:
                    inputDataFilters_Table = InputDataFilters_Tables.SUPERDETTAGLI;
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
            var filters = Context.Applicablefilters.Where(af => af.Table == inputDataFilters_Table && af.SelectedValues.Any()).ToList();
            #endregion

            // Determina l'ultima riga con dati nel foglio sorgente
            int lastRowSource = wsSource.Dimension.End.Row;

            // mi posizione sulla riga degli headers cosi al primo passaggio sarò già sulla prima riga con i dati
            int destRowIndex = destHeaderRow;

            // Copia dati riga per riga, rispettando i nomi delle colonne
            for (int rowSourceIndex = souceHeaderRow + 1; rowSourceIndex <= lastRowSource; rowSourceIndex++)
            {
                // avanzo di una riga
                destRowIndex++;

                #region verifico che la riga non sia da saltare per via dei filtri non corrispondenti
                if (filters.Any())
                {
                    foreach (var filter in filters)
                    {
                        // trovo la colonna sorgente usando il nome della colonna di destinazione
                        int sourceColumnIndex = sourceHeaders[filter.FieldName.ToLower()];

                        // prendo il valore dalla sorgente
                        var value = wsSource.Cells[rowSourceIndex, sourceColumnIndex].Value;

                        // se il valore (non null) non è presente tra i valori selezionati, la riga viene saltata
                        if (value != null && !filter.SelectedValues.Any(_ => _.Equals(value)))
                        { continue; }
                    }
                }
                #endregion

                foreach (var kvp in destHeaders)
                {
                    string destHeader = kvp.Key;
                    int destColumnIndex = kvp.Value;

                    // posso risparmiarmi questo check
                    //// Se il foglio sorgente ha la stessa colonna, copia il valore
                    //if (sourceHeaders.ContainsKey(destHeader))
                    //{

                    // trovo la colonna sorgente usando il nome della colonna di destinazione
                    int sourceColumnIndex = sourceHeaders[destHeader];

                    // prendo il valore dalla sorgente
                    var value = wsSource.Cells[rowSourceIndex, sourceColumnIndex].Value;



                    // lo scrivo nella destinazione
                    wsDest.Cells[destRowIndex, destColumnIndex].Value = value;
                    //}
                }
            }

            //todo: allungare tabella con formule

            // Salva le modifiche
            packageDest.Save();
        }
    }
}