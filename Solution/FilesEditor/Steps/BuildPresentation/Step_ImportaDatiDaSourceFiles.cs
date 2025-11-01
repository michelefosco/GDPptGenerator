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
    internal class Step_ImportaDatiDaSourceFiles : StepBase
    {
        public Step_ImportaDatiDaSourceFiles(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            Context.DebugInfoLogger.LogStepContext("Step_ImportaDatiDaSourceFiles", Context);

            ImportaSourceFile(
                    sourceFileType: FileTypes.Budget,
                    sourceFilePath: Context.FileBudgetPath,
                    sourceWorksheetName: WorksheetNames.SOURCEFILE_BUDGET_DATA,
                    souceHeadersRow: Context.Configurazione.INPUT_FILES_BUDGET_HEADERS_ROW,
                    sourceHeadersFirstColumn: Context.Configurazione.INPUT_FILES_BUDGET_HEADERS_FIRST_COL,
                    //
                    destWorksheetName: WorksheetNames.DATASOURCE_BUDGET_DATA,
                    destHeadersRow: Context.Configurazione.DATASOURCE_BUDGET_HEADERS_ROW,
                    destHeadersFirstColumn: Context.Configurazione.DATASOURCE_BUDGET_HEADERS_FIRST_COL
                );

            ImportaSourceFile(
                    sourceFileType: FileTypes.Forecast,
                    sourceFilePath: Context.FileForecastPath,
                    sourceWorksheetName: WorksheetNames.SOURCEFILE_FORECAST_DATA,
                    souceHeadersRow: Context.Configurazione.INPUT_FILES_FORECAST_HEADERS_ROW,
                    sourceHeadersFirstColumn: Context.Configurazione.INPUT_FILES_FORECAST_HEADERS_FIRST_COL,
                    //
                    destWorksheetName: WorksheetNames.DATASOURCE_FORECAST_DATA,
                    destHeadersRow: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_ROW,
                    destHeadersFirstColumn: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_FIRST_COL
                );

            //ImportaSourceFile(
            //        sourceFileType: FileTypes.RunRate,
            //        sourceFilePath: Context.FileRunRatePath,
            //        sourceWorksheetName: WorksheetNames.SOURCEFILE_RUN_RATE_DATA,
            //        souceHeadersRow: Context.Configurazione.INPUT_FILES_RUNRATE_HEADERS_ROW,
            //        sourceHeadersFirstColumn: Context.Configurazione.INPUT_FILES_RUNRATE_HEADERS_FIRST_COL,
            //        //
            //        destWorksheetName: WorksheetNames.DATASOURCE_RUN_RATE_DATA,
            //        destHeadersRow: Context.Configurazione.INPUT_FILES_RUNRATE_HEADERS_ROW,
            //        destHeadersFirstColumn: Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_FIRST_COL
            //    );

            ImportaSourceFile(
                    sourceFileType: FileTypes.SuperDettagli,
                    sourceFilePath: Context.FileSuperDettagliPath,
                    sourceWorksheetName: WorksheetNames.SOURCEFILE_SUPERDETTAGLI_DATA,
                    souceHeadersRow: Context.Configurazione.INPUT_FILES_SUPERDETTAGLI_HEADERS_ROW,
                    sourceHeadersFirstColumn: Context.Configurazione.INPUT_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL,
                    //
                    destWorksheetName: WorksheetNames.DATASOURCE_SUPERDETTAGLI_DATA,
                    destHeadersRow: Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_ROW,
                    destHeadersFirstColumn: Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL
                );

            // todo: salvare subito o lasciare tutto l'operazione finale?
            //Context.EpplusHelperDataSource.Save();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void ImportaSourceFile(
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
            // variabili per il conteggio delle righe Eliminate, Aggiunge e Preservate (quando in modalità append su Superdettagli)
            int totRighePreservate = 0;
            int totRigheEliminate = 0;
            int totRigheAggiunte = 0;

            #region WorkSheets sorgente e destinazione
            // Foglio sorgente
            var packageSource = new ExcelPackage(new FileInfo(sourceFilePath));
            var worksheetSource = packageSource.Workbook.Worksheets[sourceWorksheetName];

            // Foglio destinazione
            var worksheetDest = Context.EpplusHelperDataSource.ExcelPackage.Workbook.Worksheets[destWorksheetName];
            #endregion


            #region Lettura degli headers del foglio sorgente e foglio destinazione
            // headers del folio sorgente (Budget, Forecast, RunRate, SuperDettagli)
            var sourceHeadersDictionary = GetHeadersDictionary(workSheet: worksheetSource,
                                        headersRow: souceHeadersRow,
                                        headerFirstColumn: sourceHeadersFirstColumn);
            // headers del golio destinazione (Datasource)
            var destHeadersDictionary = GetHeadersDictionary(workSheet: worksheetDest,
                            headersRow: destHeadersRow,
                            headerFirstColumn: destHeadersFirstColumn);
            #endregion


            #region Verifico che tutti gli headers presenti nella destinazione siano effettivamente presenti nella sorgente
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


            #region Preparo una struttura più snella che contenga le informazioni su filtri
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


            #region Se in modalità append su SuperDettagli preservo le righe il cui valore nella colonna "anno" è diverso da quello scelto come periodo
            if (sourceFileType == FileTypes.SuperDettagli && Context.AppendCurrentYear_FileSuperDettagli)
            {
                #region Determino l'indice della colonna "anno"
                if (!destHeadersDictionary.ContainsKey("anno"))
                { throw new Exception("The sheet “Superdettagli” does not contain the required column “year” needed to handle the data append."); }
                var colonnaAnnoIndex = destHeadersDictionary["anno"];
                #endregion


                #region Controllo che la sorgente non contenga valori di anno diversi da quello del periodo. In tal caso sollevo un warning
                for (var rowIndex = souceHeadersRow + 1; rowIndex <= worksheetSource.Dimension.End.Row; rowIndex++)
                {
                    // lettura del valore "anno"
                    if (!int.TryParse(worksheetSource.Cells[rowIndex, colonnaAnnoIndex].Value.ToString(), out int anno))
                    { throw new Exception($"The row {rowIndex} in the 'anno' column of the input file 'Superdettagli' does not contain an integer number."); }

                    if (anno != Context.PeriodYear)
                    {
                        Context.AddWarning($"The input file 'Super dettagli' contains at least one year that is different from the year selected as the period date ({Context.PeriodYear}).");
                        break;
                    }
                }
                #endregion



   
                // scorro le righe esistenti valutanto il valore delle celle della colonna "anno"
                for (int rowIndex = destHeadersRow + 1; rowIndex <= worksheetDest.Dimension.Rows; rowIndex++)
                {
                    // lettura del valore "anno"
                    if (!int.TryParse(worksheetDest.Cells[rowIndex, colonnaAnnoIndex].Value.ToString(), out int anno))
                    { throw new Exception($"The row {rowIndex} in the 'anno' column of the 'Superdettagli' sheet does not contain an integer number."); }

                    // elimino le righe il cui valore nella cella "anno" è uguale a Context.PeriodYear
                    if (anno == Context.PeriodYear)
                    {
                        worksheetDest.DeleteRow(rowIndex, 1, true);
                        totRigheEliminate++;
                        rowIndex--;
                    }
                    else
                    {
                        totRighePreservate++;
                    }
                }
            }
            #endregion


            // Per l'aggiunta delle righe parto sempre dalla prima immediatamente dopo gli headers per asicurarmi di preservare le formule inserendo nuove righe
            // Rappresenta la riga del foglio di destinazione in cui scrivere la prossima riga
            var destRowIndex = destHeadersRow + 1;



            #region Scorro tutte le righe della sorgente a partire da quella immediatamente successiva alla riga con gli headers
            for (var rowSourceIndex = souceHeadersRow + 1; rowSourceIndex <= worksheetSource.Dimension.End.Row; rowSourceIndex++)
            {
                #region Verifico che la riga non sia da saltare per via dei filtri non corrispondenti
                bool skippaRigaPerFiltro = false;
                if (filters.Any())
                {
                    foreach (var filter in filters)
                    {
                        if (!sourceHeadersDictionary.ContainsKey(filter.FieldName.ToLower()))
                        {
                            throw new ArgumentOutOfRangeException($"One of the filters set does not have a corresponding header in the source file");
                        }

                        // trovo la colonna sorgente usando il nome della colonna di destinazione
                        int sourceColumnIndex = sourceHeadersDictionary[filter.FieldName.ToLower()];

                        // prendo il valore dalla sorgente
                        var value = worksheetSource.Cells[rowSourceIndex, sourceColumnIndex].Value;

                        // se il valore (non null) non è presente tra i valori selezionati, la riga viene saltata
                        if (value != null && !filter.SelectedValues.Any(_ => _.Equals(value)))
                        {
                            skippaRigaPerFiltro = true;
                            break;
                        }
                    }
                }
                if (skippaRigaPerFiltro)
                { continue; }
                #endregion

                #region Allungo la tabella di un riga in modo da conservare le formule
                if (sourceFileType != FileTypes.RunRate)
                {
                    worksheetDest.InsertRow(destRowIndex, 1);
                    totRigheAggiunte++;
                }
                #endregion

                #region Copio i valori per le colonne presenti nel foglio di destinazione (uso il dizionario "destHeadersDictionary")
                foreach (var kvp in destHeadersDictionary)
                {
                    // determino l'indice della colonna sul foglio di destinazione
                    var destHeader = kvp.Key;
                    var destColumnIndex = kvp.Value;

                    // determino l'indice della colonnasul foglio sorgente
                    var sourceColumnIndex = sourceHeadersDictionary[destHeader];

                    // prendo il valore dalla cella del foglio sorgente
                    var valueFromSource = worksheetSource.Cells[rowSourceIndex, sourceColumnIndex].Value;

                    // Vengono valutate eventuali sostituzioni di valore in base agli aliases dichiarati per i campi "Business TMP" e "Categoria" dei file "Budget" e "Forecast"
                    if (sourceFileType == FileTypes.Budget || sourceFileType == FileTypes.Forecast)
                    { valueFromSource = ApplicaAliasToValue(destHeader, valueFromSource.ToString()); }

                    // Scrivo il valore nella cella di destinazione
                    worksheetDest.Cells[destRowIndex, destColumnIndex].Value = valueFromSource;
                }
                #endregion

                // avanzo di una riga
                if (sourceFileType != FileTypes.RunRate)
                { destRowIndex++; }
            }
            #endregion

            // ritorno indietro di uno per mantenere il significato del valore (Rappresenta la riga del foglio di destinazione in cui scrivere la prossima riga) 
            destRowIndex--;


            #region Cancellazione delle righe in più, ovvero quelle gia esistenti ma non più utilizzate (esempio ho meno righe dell'aggiornamento precedente)
            if (sourceFileType != FileTypes.RunRate)
            {
                // aggiungo al conteggio le righe preservate, in modo da non canellarle
                destRowIndex += totRighePreservate;

                // la cancellazione deve avvenire dall'ultima riga indietro in quanto le righe eliminate shiftano verso il basso e gli indici delle righe vengono aggiornati
                for (int rowIndex = worksheetDest.Dimension.Rows; rowIndex > destRowIndex; rowIndex--)
                {
                    worksheetDest.DeleteRow(rowIndex, 1, true);
                    totRigheEliminate++;
                }
            }
            #endregion


            #region Per la modalità "Append" della tabella "Superdettagli" è necessario ordinare la tabella
            // per il campo "anno" in quanto, per preservare le formule, le nuove righe sono state aggiunte immediatamente dopo la riga degli headers
            // e quindi precedentemente alle righe già esistenti che hanno verosimilmente un numero di anno inferiore
            if (sourceFileType == FileTypes.SuperDettagli && Context.AppendCurrentYear_FileSuperDettagli)
            {
                var colonnaAnnoPositionZeroBased = destHeadersDictionary["anno"] - 1;
                Context.EpplusHelperDataSource.OrdinaTabella(destWorksheetName, destHeadersRow + 1, 1, worksheetDest.Dimension.End.Row, worksheetDest.Dimension.End.Column, colonnaAnnoPositionZeroBased);
            }
            #endregion

            // Log delle informazioni
            Context.DebugInfoLogger.LogRigheSourceFiles(sourceFileType, totRighePreservate, totRigheEliminate, totRigheAggiunte);
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