using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using OfficeOpenXml;
using System;
using System.Configuration;
using System.IO;
using System.Linq;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_ImportaDati_SuperDettagli : StepBase
    {
        public override string StepName => "Step_ImportaDati_SuperDettagli";

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
        public Step_ImportaDati_SuperDettagli(StepContext context) : base(context)
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

            int newRowsYear = Context.PeriodYear;

           // var superDettagliEPPlusHelper = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(Context.FileSuperDettagliPath, FileTypes.SuperDettagli);

            #region Preparo una struttura più snella che contenga le informazioni su filtri
            var filters = Context.ApplicableFilters.Where(_ => _.Table == InputDataFilters_Tables.SUPERDETTAGLI
                                                            && _.SelectedValues.Any())
                                                    .ToList();
            #endregion


            #region Lettura degli headers del foglio sorgente e foglio destinazione
            var sourceHeaders = Context.SuperdettagliFileEPPlusHelper.GetHeadersFromRow(sourceWorksheetName, souceHeadersRow, sourceHeadersFirstColumn, true).Select(_ => _.ToLower()).ToList();
            var destHeaders = Context.DataSourceEPPlusHelper.GetHeadersFromRow(destWorksheetName, destHeadersRow, destHeadersFirstColumn, true).Select(_ => _.ToLower()).ToList();
            #endregion


            #region WorkSheets sorgente e destinazione
            // Foglio sorgente
            var worksheetSource = Context.SuperdettagliFileEPPlusHelper.ExcelPackage.Workbook.Worksheets[sourceWorksheetName];

            // Foglio destinazione
            var worksheetDest = Context.DataSourceEPPlusHelper.ExcelPackage.Workbook.Worksheets[destWorksheetName];
            #endregion


            #region Se in modalità append su SuperDettagli preservo le righe il cui valore nella colonna "anno" è diverso da quello scelto come periodo
            if (Context.AppendCurrentYear_FileSuperDettagli)
            {
                // scorro le righe esistenti valutando il valore delle celle della colonna "anno"
                for (int rowIndex = destHeadersRow + 1; rowIndex <= worksheetDest.Dimension.Rows; rowIndex++)
                {
                    // lettura del valore "anno"
                    if (!int.TryParse(worksheetDest.Cells[rowIndex, Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_YEAR_COL].Value.ToString(), out int annoRigaCorrente))
                    { throw new Exception($"The row {rowIndex} in the 'anno' column of the 'Superdettagli' sheet does not contain an integer number."); }

                    // elimino le righe il cui valore nella cella "anno" è uguale a Context.PeriodYear
                    if (annoRigaCorrente == newRowsYear)
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


            #region Sfoltisco le righe della sorgente in base ai filtri impostati
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

                    for (var rowSourceIndex = souceHeadersRow + 1; rowSourceIndex <= worksheetSource.Dimension.End.Row; rowSourceIndex++)
                    {

                        // prendo il valore dalla sorgente
                        var value = worksheetSource.Cells[rowSourceIndex, sourceColumnIndex].Value;

                        // se il valore (non null) non è presente tra i valori selezionati, la riga viene saltata
                        if (value != null && !filter.SelectedValues.Any(_ => _.Equals(value)))
                        {
                            worksheetSource.DeleteRow(rowSourceIndex, 1, true);
                            continue;
                        }
                    }
                }
            }
            #endregion

            var numeroDiColonneDaCopiare = destHeaders.Count;
            var righeDaAggiungere = worksheetSource.Dimension.End.Row - souceHeadersRow;

            // Per l'aggiunta delle righe parto sempre dalla prima immediatamente dopo gli headers per asicurarmi di preservare le formule inserendo nuove righe
            // Rappresenta la riga del foglio di destinazione in cui scrivere la prossima riga
            var destRowIndex = destHeadersRow + 1;

            #region Allungo la tabella
            worksheetDest.InsertRow(destRowIndex, righeDaAggiungere);
            totRigheAggiunte += righeDaAggiungere;
            #endregion


            #region Copio i valori per le colonne presenti nel foglio di destinazione (uso il dizionario "destHeadersDictionary")
            // Range sorgente
            var sourceRange = worksheetSource.Cells[
                    souceHeadersRow + 1,                                                                            // row start
                    Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL,                            // col start 
                    worksheetSource.Dimension.End.Row,                                                              // row end
                    Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL + numeroDiColonneDaCopiare  // col end
                    ];

            // Incollo nel range destinazione
            worksheetDest.Cells[
                    destRowIndex,                                                                                   // row start
                    Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL,                              // col start 
                    destRowIndex + righeDaAggiungere - 1,                                                           // row end
                    Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL + numeroDiColonneDaCopiare    // col end
                    ].Value = sourceRange.Value;

            // Incollo il valore dell'anno
            worksheetDest.Cells[
                    destRowIndex,                           // row start
                    Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_YEAR_COL, // col "year"
                    destRowIndex + righeDaAggiungere - 1,       // row end
                    Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_YEAR_COL  // col "year"
                    ].Value = newRowsYear;
            #endregion


            #region Cancellazione delle righe in più, ovvero quelle gia esistenti ma non più utilizzate (esempio ho meno righe dell'aggiornamento precedente)
            // determino l'ultima riga da preservare
            var ultimaRigaDaMantenere = destHeadersRow + totRigheAggiunte + totRighePreservate;

            // la cancellazione deve avvenire dall'ultima riga indietro in quanto le righe eliminate shiftano verso il basso e gli indici delle righe vengono aggiornati
            for (int rowIndex = worksheetDest.Dimension.Rows; rowIndex > ultimaRigaDaMantenere; rowIndex--)
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
                var colonnaAnnoPositionZeroBased = Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_YEAR_COL;
                Context.DataSourceEPPlusHelper.OrdinaTabella(destWorksheetName, destHeadersRow + 1, 1, worksheetDest.Dimension.End.Row, worksheetDest.Dimension.End.Column, colonnaAnnoPositionZeroBased);
            }
            #endregion


            #region Eccezione nel caso in cui, dopo l'applicazione dei filtri, non ci siano righe da importare
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
            #endregion

            // Log delle informazioni
            Context.DebugInfoLogger.LogRigheSourceFiles(FileTypes.SuperDettagli, totRighePreservate, totRigheEliminate, totRigheAggiunte);

           // superDettagliEPPlusHelper.Close();
        }
    }
}