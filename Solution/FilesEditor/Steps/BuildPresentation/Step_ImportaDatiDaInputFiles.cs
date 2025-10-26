using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FilesEditor.Steps.BuildPresentation
{
    internal class Step_ImportaDatiDaInputFiles : StepBase
    {
        public Step_ImportaDatiDaInputFiles(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            Context.DebugInfoLogger.LogStepContext("Step_ImportaDatiDaInputFiles", Context);

            leggiInputFile(
                    sourceFileType: FileTypes.Budget,
                    sourceFilePath: Context.FileBudgetPath,
                    sourceWorksheetName: WorksheetNames.INPUTFILES_BUDGET_DATA,
                    souceHeadersRow: Context.Configurazione.INPUT_FILES_BUDGET_HEADERS_ROW,
                    sourceHeadersFirstColumn: Context.Configurazione.INPUT_FILES_BUDGET_HEADERS_FIRST_COL,
                    //
                    destWorksheetName: WorksheetNames.DATASOURCE_BUDGET_DATA,
                    destHeadersRow: Context.Configurazione.DATASOURCE_BUDGET_HEADERS_ROW,
                    destHeadersFirstColumn: Context.Configurazione.DATASOURCE_BUDGET_HEADERS_FIRST_COL
                );

            leggiInputFile(
                    sourceFileType: FileTypes.Forecast,
                    sourceFilePath: Context.FileForecastPath,
                    sourceWorksheetName: WorksheetNames.INPUTFILES_FORECAST_DATA,
                    souceHeadersRow: Context.Configurazione.INPUT_FILES_FORECAST_HEADERS_ROW,
                    sourceHeadersFirstColumn: Context.Configurazione.INPUT_FILES_FORECAST_HEADERS_FIRST_COL,
                    //
                    destWorksheetName: WorksheetNames.DATASOURCE_FORECAST_DATA,
                    destHeadersRow: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_ROW,
                    destHeadersFirstColumn: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_FIRST_COL
                );

            leggiInputFile(
                    sourceFileType: FileTypes.RunRate,
                    sourceFilePath: Context.FileRunRatePath,
                    sourceWorksheetName: WorksheetNames.INPUTFILES_RUN_RATE_DATA,
                    souceHeadersRow: Context.Configurazione.INPUT_FILES_RUNRATE_HEADERS_ROW,
                    sourceHeadersFirstColumn: Context.Configurazione.INPUT_FILES_RUNRATE_HEADERS_FIRST_COL,
                    //
                    destWorksheetName: WorksheetNames.DATASOURCE_RUN_RATE_DATA,
                    destHeadersRow: Context.Configurazione.INPUT_FILES_RUNRATE_HEADERS_ROW,
                    destHeadersFirstColumn: Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_FIRST_COL
                );

            //todo: gestione particolare per Superdettagli
            leggiInputFile(
                    sourceFileType: FileTypes.SuperDettagli,
                    sourceFilePath: Context.FileSuperDettagliPath,
                    sourceWorksheetName: WorksheetNames.INPUTFILES_SUPERDETTAGLI_DATA,
                    souceHeadersRow: Context.Configurazione.INPUT_FILES_SUPERDETTAGLI_HEADERS_ROW,
                    sourceHeadersFirstColumn: Context.Configurazione.INPUT_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL,
                    //
                    destWorksheetName: WorksheetNames.DATASOURCE_SUPERDETTAGLI_DATA,
                    destHeadersRow: Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_ROW,
                    destHeadersFirstColumn: Context.Configurazione.INPUT_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL
                );

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void leggiInputFile(
                FileTypes sourceFileType,
                string sourceFilePath,
                string sourceWorksheetName,
                int souceHeadersRow,
                int sourceHeadersFirstColumn,
                //
                string destWorksheetName,
                int destHeadersRow,
                int destHeadersFirstColumn
            )
        {
            // Foglio sorgente
            var packageSource = new ExcelPackage(new FileInfo(sourceFilePath));
            var worksheetSource = packageSource.Workbook.Worksheets[sourceWorksheetName];

            // Foglio destinazione
            var worksheetDest = Context.EpplusHelperDataSource.ExcelPackage.Workbook.Worksheets[destWorksheetName];

            #region Lettura degli headers del foglio sorgente e foglio destinazione
            var sourceHeadersDictionary = GetHeadersDictionary(workSheet: worksheetSource,
                                        headersRow: souceHeadersRow,
                                        headerFirstColumn: sourceHeadersFirstColumn);

            // Lettura headers del foglio di destinazione
            //var destHeadersDictionary = new Dictionary<string, int>();
            //var destCol = destHeadersFirstColumn;
            //while (wsDest.Cells[destHeadersRow, destCol].Value != null)
            //{
            //    string header = wsDest.Cells[destHeadersRow, destCol].Text.Trim().ToLower();
            //    destHeadersDictionary[header] = destCol;
            //    destCol++;
            //}
            var destHeadersDictionary = GetHeadersDictionary(workSheet: worksheetDest,
                            headersRow: destHeadersRow,
                            headerFirstColumn: destHeadersFirstColumn);
            #endregion

            #region Verifico che tutti gli headers presenti nella destinazione sia effettivamente presenti nella sorgente
            foreach (var kvp in destHeadersDictionary)
            {
                string destHeader = kvp.Key;
                // Se il foglio sorgente non ha tutti gli headers necessari sollevo l'eccezione
                if (!sourceHeadersDictionary.ContainsKey(destHeader))
                {
                    throw new ManagedException(
                        filePath: sourceFilePath,
                        fileType: sourceFileType,
                        //
                        worksheetName: sourceWorksheetName,
                        cellRow: souceHeadersRow,
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
            var filters = Context.ApplicableFilters.Where(af => af.Table == inputDataFilters_Table && af.SelectedValues.Any()).ToList();
            #endregion

            // Determina l'ultima riga con dati nel foglio sorgente
            var lastRowSource = worksheetSource.Dimension.End.Row;

            // mi posizione sulla riga immediatamente dopo quella degli headers della destinazione
            var destRowIndex = destHeadersRow + 1;

            var table = worksheetDest.Tables[0];
            var tableAdress = table.Address;

            // Copia dati riga per riga, rispettando i nomi delle colonne
            for (var rowSourceIndex = souceHeadersRow + 1; rowSourceIndex <= lastRowSource; rowSourceIndex++)
            {
                #region verifico che la riga non sia da saltare per via dei filtri non corrispondenti
                bool skippaRiga = false;
                if (filters.Any())
                {
                    foreach (var filter in filters)
                    {
                        // trovo la colonna sorgente usando il nome della colonna di destinazione
                        int sourceColumnIndex = sourceHeadersDictionary[filter.FieldName.ToLower()];

                        // prendo il valore dalla sorgente
                        var value = worksheetSource.Cells[rowSourceIndex, sourceColumnIndex].Value;

                        // se il valore (non null) non è presente tra i valori selezionati, la riga viene saltata
                        if (value != null && !filter.SelectedValues.Any(_ => _.Equals(value)))
                        {
                            skippaRiga = true;
                            break;
                        }
                    }
                }
                if (skippaRiga)
                { continue; }
                #endregion

                // Allungo la tabella di un riga in modo da conservare le formule
                worksheetDest.InsertRow(destRowIndex, 1);

                foreach (var kvp in destHeadersDictionary)
                {
                    var destHeader = kvp.Key;
                    var destColumnIndex = kvp.Value;

                    // trovo la colonna sorgente usando il nome della colonna di destinazione
                    var sourceColumnIndex = sourceHeadersDictionary[destHeader];

                    // prendo il valore dalla sorgente
                    var valueFromSource = worksheetSource.Cells[rowSourceIndex, sourceColumnIndex].Value;

                    // Vengono valutati gli aliases per i campi "Business TMP" e "Categoria" dei file "Budget" e "Forecast"
                    if (sourceFileType == FileTypes.Budget || sourceFileType == FileTypes.Forecast)
                    { valueFromSource = ApplicaAliasToValue(destHeader, valueFromSource.ToString()); }

                    // lo scrivo nella destinazione
                    worksheetDest.Cells[destRowIndex, destColumnIndex].Value = valueFromSource;
                }

                // avanzo di una riga
                destRowIndex++;
            }

            // cancello le righe inutilizzate
            worksheetDest.DeleteRow(destRowIndex, worksheetDest.Dimension.End.Row, true);

            // todo: salvare subito o lasciare tutto l'operazione finale?
            Context.EpplusHelperDataSource.ExcelPackage.Save();
        }

        Dictionary<string, int> GetHeadersDictionary(ExcelWorksheet workSheet, int headersRow, int headerFirstColumn)
        {
            var sourceHeaders = new Dictionary<string, int>();
            var sourceCol = headerFirstColumn;
            while (workSheet.Cells[headersRow, sourceCol].Value != null)
            {
                var header = workSheet.Cells[headersRow, sourceCol].Text.Trim().ToLower();
                sourceHeaders[header] = sourceCol;
                sourceCol++;
            }

            return sourceHeaders;
        }

        private string ApplicaAliasToValue(string header, string value)
        {
            if (string.IsNullOrEmpty(header))
            { return value; }

            #region Controllo se il valore appartiene ad uno di quelli interessati da Aliases
            List<AliasDefinition> aliasesToCheck = null;
            if (header.Equals(Values.ALIAS_HEADER_BUSINESS_TMP, StringComparison.InvariantCultureIgnoreCase))
            {
                aliasesToCheck = Context.AliasDefinitions_BusinessTMP;
            }
            else if (header.Equals(Values.ALIAS_HEADER_CATEGORIA, StringComparison.InvariantCultureIgnoreCase))
            {
                aliasesToCheck = Context.AliasDefinitions_Categoria;
            }
            else
            {
                // non è necessario applicare gli alias
                return value;
            }
            #endregion

            // non ci sono aliase da verificare
            if (aliasesToCheck == null || aliasesToCheck.Count == 0)
            { return value; }

            #region Cerco un alias che corrisponda
            // Controllo prima gli aliases fissi (senza regular expressions)
            foreach (AliasDefinition alias in aliasesToCheck.Where(_ => !_.IsRegularExpression))
            {
                if (alias.RawValue.Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase))
                { return alias.NewValue; }
            }

            // Controllo successivamente gli aliases con regular expressions
            foreach (AliasDefinition alias in aliasesToCheck.Where(_ => _.IsRegularExpression))
            {
                if (ValuesHelper.StringMatch(value, alias.RawValue))
                { return alias.NewValue; }
            }
            #endregion

            // Nessun alias applicato
            return value;
        }
    }
}