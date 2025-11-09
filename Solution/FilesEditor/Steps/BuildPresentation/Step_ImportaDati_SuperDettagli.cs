using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_ImportaDati_SuperDettagli_2 : StepBase
    {
        public override string StepName => "Step_ImportaDati_SuperDettagli_2";

        internal override void BeforeTask()
        {
            Context.DebugInfoLogger.LogStepContext(StepName, Context);
        }

        internal override void ManageInfoAboutPerformedStepTask(TimeSpan timeSpent)
        {
            Context.DebugInfoLogger.LogPerformance(StepName, timeSpent);
        }

        internal override void AfterTask()
        {
            Context.DebugInfoLogger.LogStepContext(StepName, Context);
        }
        public Step_ImportaDati_SuperDettagli_2(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            ImportaSourceFile(
                    //sourceFileType: FileTypes.SuperDettagli,
                    sourceFilePath: Context.FileSuperDettagliPath,
                    sourceWorksheetName: WorksheetNames.SOURCEFILE_SUPERDETTAGLI_DATA,
                    souceHeadersRow: Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_ROW,
                    sourceHeadersFirstColumn: Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL,
                    //
                    destWorksheetName: WorksheetNames.DATASOURCE_SUPERDETTAGLI_DATA,
                    destHeadersRow: Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_ROW,
                    destHeadersFirstColumn: Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL
                );

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void ImportaSourceFile(
                // FileTypes sourceFileType,
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




            //#region Lettura degli headers del foglio sorgente e foglio destinazione
            //// headers del folio sorgente (Budget, Forecast, SuperDettagli)
            //var sourceHeaders = GetHeadersList(workSheet: worksheetSource,
            //                                headersRow: souceHeadersRow,
            //                                headerFirstColumn: sourceHeadersFirstColumn);
            //// headers del golio destinazione (Datasource)
            //var destHeaders = GetHeadersList(workSheet: worksheetDest,
            //                                headersRow: destHeadersRow,
            //                                headerFirstColumn: destHeadersFirstColumn);
            //#endregion


            var superDettagliEPPlusHelper = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(Context.FileSuperDettagliPath, FileTypes.SuperDettagli);


            #region Lettura degli headers del foglio sorgente e foglio destinazione
            var sourceHeaders = superDettagliEPPlusHelper.GetHeadersFromRow(sourceWorksheetName, souceHeadersRow, sourceHeadersFirstColumn, true);
            var destHeaders = Context.EpplusHelperDataSource.GetHeadersFromRow(destWorksheetName, destHeadersRow, destHeadersFirstColumn, true);
            #endregion

            // IMPORTANTE: La validazione degli headers avviene all'interno dello step 'Step_ValidazioniPreliminari_SuperDettagli'
            #region Verifico che tutti gli headers necessari per la destinazione siano presenti nella sorgente (e nella giusta posizione)
            //if (sourceHeaders.Count < destHeaders.Count)
            //{
            //    throw new ManagedException(
            //            filePath: sourceFilePath,
            //            fileType: FileTypes.SuperDettagli,
            //            //
            //            worksheetName: sourceWorksheetName,
            //            cellRow: null,
            //            cellColumn: null,
            //            valueHeader: ValueHeaders.None,
            //            value: null,
            //            //
            //            errorType: ErrorTypes.MissingValue,
            //            userMessage: "There are fewer headers in the Superdettagli file than required to complete the DataSource file.\nAll headers in the DataSource file(worksheet Superdettagli) must also be present in the Superdettagli file, in the same order."
            //            );
            //}

            //for (int j = 0; j < destHeaders.Count; j++)
            //{
            //    if (!sourceHeaders[j].Equals(destHeaders[j], StringComparison.InvariantCultureIgnoreCase))
            //    {
            //        throw new ManagedException(
            //            filePath: sourceFilePath,
            //            fileType: FileTypes.SuperDettagli,
            //            //
            //            worksheetName: sourceWorksheetName,
            //            cellRow: souceHeadersRow,
            //            cellColumn: null,
            //            valueHeader: ValueHeaders.None,
            //            value: destHeaders[j],
            //            //
            //            errorType: ErrorTypes.MissingValue,
            //            userMessage: $"The header '{destHeaders[j]}' required for the file Datasource is missing (or located in the wrong position) in the file 'Superdettagli'.\nAll headers in the DataSource file (worksheet Superdettagli) must also be present in the Superdettagli file, in the same order."
            //            );
            //    }
            //}
            #endregion


            #region Preparo una struttura più snella che contenga le informazioni su filtri
            var filters = Context.ApplicableFilters.Where(_ => _.Table == InputDataFilters_Tables.SUPERDETTAGLI
                                                            && _.SelectedValues.Any())
                                                    .ToList();
            #endregion


            #region Se in modalità append su SuperDettagli preservo le righe il cui valore nella colonna "anno" è diverso da quello scelto come periodo
            if (Context.AppendCurrentYear_FileSuperDettagli)
            {
                #region Determino l'indice della colonna "anno"
                if (!destHeaders.Any(_ => _.Equals("anno", StringComparison.InvariantCultureIgnoreCase)))
                { throw new Exception("The sheet “Superdettagli” does not contain the required column “year” needed to handle the data append."); }
                var sourceColonnaAnnoIndex = Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL + destHeaders.IndexOf("anno");
                var destColonnaAnnoIndex = Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL + sourceHeaders.IndexOf("anno");
                #endregion


                #region Controllo che la sorgente non contenga valori di anno diversi da quello del periodo. In tal caso sollevo un warning
                for (var rowIndex = souceHeadersRow + 1; rowIndex <= worksheetSource.Dimension.End.Row; rowIndex++)
                {
                    // lettura del valore "anno"
                    if (!int.TryParse(worksheetSource.Cells[rowIndex, sourceColonnaAnnoIndex].Value.ToString(), out int anno))
                    { throw new Exception($"The row {rowIndex} in the 'anno' column of the input file 'Superdettagli' does not contain an integer number."); }

                    if (anno != Context.PeriodYear)
                    {
                        Context.AddWarning($"The input file 'Super dettagli' contains at least one year that is different from the year selected as the period date ({Context.PeriodYear}).");
                        break;
                    }
                }
                #endregion


                // scorro le righe esistenti valutando il valore delle celle della colonna "anno"
                for (int rowIndex = destHeadersRow + 1; rowIndex <= worksheetDest.Dimension.Rows; rowIndex++)
                {
                    // lettura del valore "anno"
                    if (!int.TryParse(worksheetDest.Cells[rowIndex, destColonnaAnnoIndex].Value.ToString(), out int anno))
                    { throw new Exception($"The row {rowIndex} in the 'anno' column of the 'Superdettagli' sheet does not contain an integer number."); }

                    // elimino le righe il cui valore nella cella "anno" è uguale a Context.PeriodYear
                    if (anno == Context.PeriodYear)
                    {
                        worksheetDest.DeleteRow(rowIndex, 1, true);
                        totRigheEliminate++;
                        rowIndex--; // rimango sulla stessa riga
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


            int numeroHeaders = destHeaders.Count;
            #region Scorro tutte le righe della sorgente a partire da quella immediatamente successiva alla riga con gli headers
            for (var rowSourceIndex = souceHeadersRow + 1; rowSourceIndex <= worksheetSource.Dimension.End.Row; rowSourceIndex++)
            {
                #region Verifico che la riga non sia da saltare per via dei filtri non corrispondenti
                bool skippaRigaPerFiltro = false;
                if (filters.Any())
                {
                    foreach (var filter in filters)
                    {
                        if (!sourceHeaders.Any(_ => _.Equals(filter.FieldName, StringComparison.InvariantCultureIgnoreCase)))
                        {
                            throw new ArgumentOutOfRangeException($"One of the filters set does not have a corresponding header in the source file");
                        }

                        // trovo la colonna sorgente usando il nome della colonna di destinazione
                        int sourceColumnIndex = sourceHeaders.IndexOf(filter.FieldName.ToLower()) + 1;

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
                worksheetDest.InsertRow(destRowIndex, 1);
                totRigheAggiunte++;
                #endregion

                #region Copio i valori per le colonne presenti nel foglio di destinazione (uso il dizionario "destHeadersDictionary")
                // Range sorgente
                var sourceRange = worksheetSource.Cells[
                        rowSourceIndex, // row start
                        Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL,    // col start 
                        rowSourceIndex, // row end
                        Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL + numeroHeaders // col end
                        ];

                // Incollo nel range destinazione
                worksheetDest.Cells[
                        destRowIndex,   // row start
                        Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL,  // col start 
                        destRowIndex,   // row end
                        Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL + numeroHeaders   // col end
                        ].Value = sourceRange.Value;
                #endregion

                destRowIndex++;
            }
            #endregion

            // ritorno indietro di uno per mantenere il significato del valore (Rappresenta la riga del foglio di destinazione in cui scrivere la prossima riga) 
            destRowIndex--;


            #region Cancellazione delle righe in più, ovvero quelle gia esistenti ma non più utilizzate (esempio ho meno righe dell'aggiornamento precedente)
            // aggiungo al conteggio le righe preservate, in modo da non canellarle
            destRowIndex += totRighePreservate;

            // la cancellazione deve avvenire dall'ultima riga indietro in quanto le righe eliminate shiftano verso il basso e gli indici delle righe vengono aggiornati
            for (int rowIndex = worksheetDest.Dimension.Rows; rowIndex > destRowIndex; rowIndex--)
            {
                worksheetDest.DeleteRow(rowIndex, 1, true);
                totRigheEliminate++;
            }
            #endregion


            #region Per la modalità "Append" della tabella "Superdettagli" è necessario ordinare la tabella
            // per il campo "anno" in quanto, per preservare le formule, le nuove righe sono state aggiunte immediatamente dopo la riga degli headers
            // e quindi precedentemente alle righe già esistenti che hanno verosimilmente un numero di anno inferiore
            if (Context.AppendCurrentYear_FileSuperDettagli)
            {
                var colonnaAnnoPositionZeroBased = destHeaders.IndexOf("anno");
                Context.EpplusHelperDataSource.OrdinaTabella(destWorksheetName, destHeadersRow + 1, 1, worksheetDest.Dimension.End.Row, worksheetDest.Dimension.End.Column, colonnaAnnoPositionZeroBased);
            }
            #endregion

            if (totRighePreservate + totRigheAggiunte == 0)
            {
                throw new ManagedException(
                    filePath: sourceFilePath,
                    fileType: FileTypes.SuperDettagli,
                    //
                    worksheetName: sourceWorksheetName,
                    cellRow: null,
                    cellColumn: null,
                    valueHeader: ValueHeaders.None,
                    value: null,
                    //
                    errorType: ErrorTypes.NoDataAvailable,
                    userMessage: string.Format(UserErrorMessages.NoDataAvailableFromFileAfterFilters, FileTypes.SuperDettagli)
                    );
            }
            // Log delle informazioni
            Context.DebugInfoLogger.LogRigheSourceFiles(FileTypes.SuperDettagli, totRighePreservate, totRigheEliminate, totRigheAggiunte);
        }

        //List<string> GetHeadersList(ExcelWorksheet workSheet, int headersRow, int headerFirstColumn)
        //{
        //    var sourceHeaders = new List<string>();
        //    var sourceCol = headerFirstColumn;
        //    while (workSheet.Cells[headersRow, sourceCol].Value != null)
        //    {
        //        var header = workSheet.Cells[headersRow, sourceCol].Text.Trim().ToLower();
        //        sourceHeaders.Add(header);
        //        sourceCol++;
        //    }

        //    return sourceHeaders;
        //}
    }
}