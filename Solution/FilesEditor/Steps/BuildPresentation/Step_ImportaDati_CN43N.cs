using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Enums;
using OfficeOpenXml;
using System;
using System.Linq;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_ImportaDati_CN43N : StepBase
    {
        internal override string StepName => "Step_ImportaDati_CN43N";

        public Step_ImportaDati_CN43N(StepContext context) : base(context)
        { }

        //CN43N
        internal override EsitiFinali DoSpecificStepTask()
        {
            if (!string.IsNullOrEmpty(Context.FileCN43NPath))
            {
                if (Context.FileCN43_OverwriteAll)
                {
                    ImportDataFrom_CN43NFile_OverwriteAll();
                }
                else
                {
                    ImportDataFrom_CN43NFile_UpdateDuplicates();
                }
            }

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        /// <summary>
        /// Aggiornamento completo dei dati: vengono eliminate tutte le righe esistenti nel foglio di destinazione e vengono copiate tutte le righe dal foglio sorgente al foglio di destinazione
        /// </summary>
        private void ImportDataFrom_CN43NFile_OverwriteAll()
        {
            // Foglio sorgente            
            var sourceWorksheet = Context.CN43NFileEPPlusHelper.ExcelPackage.Workbook.Worksheets[1]; // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome

            // Foglio destinazione
            var destWorksheet = Context.DataSourceEPPlusHelper.ExcelPackage.Workbook.Worksheets[WorksheetNames.DATASOURCE_CN43N_DATA];
            var destHeadersRow = Context.Configurazione.DATASOURCE_CN43N_HEADERS_ROW;
            var numeroRigheIniziali = destWorksheet.Dimension.End.Row - destHeadersRow;


            #region Cleanup del foglio destinazione
            // Questo foglio (ad eccezione degli altri, non ha un oggetto "Table" per la gestione dei dati. E' sufficiente quindi eseguire il .Clear() sulle celle per eliminare le righe
            var destRangeForCleanUp = destWorksheet.Cells[
                            destHeadersRow + 1,                                         // row start,
                            Context.Configurazione.DATASOURCE_CN43N_HEADERS_FIRST_COL,  // col start
                            destWorksheet.Dimension.End.Row,                            // row end
                            destWorksheet.Dimension.End.Column                          // col end
                            ];
            destRangeForCleanUp.Clear();
            #endregion


            #region Copia e incolla del range dalla sorgente alla destinazione
            // Range sorgente
            var sourceRange = sourceWorksheet.Cells[
                            Context.Configurazione.SOURCE_FILES_CN43N_HEADERS_ROW + 1,      // row start,
                            Context.Configurazione.SOURCE_FILES_CN43N_HEADERS_FIRST_COL,    // col start
                            sourceWorksheet.Dimension.End.Row,                              // row end
                            sourceWorksheet.Dimension.End.Column                            // col end
                            ];


            // Range destinazione
            var destRange = destWorksheet.Cells[
                            destHeadersRow + 1,                                         // row start,
                            Context.Configurazione.DATASOURCE_CN43N_HEADERS_FIRST_COL,  // col start
                            sourceWorksheet.Dimension.End.Row,                          // row end
                            sourceWorksheet.Dimension.End.Column                        // col end
                            ];

            // Incollo nel range destinazione
            destRange.Value = sourceRange.Value;
            #endregion


            #region Log delle informazioni
            // Variabili per il conteggio delle righe elaborate
            var infoRowsDestinazione = new InfoRows
            {
                Iniziali = numeroRigheIniziali,
                Eliminate = numeroRigheIniziali,
                Preservate = 0,
                Riutilizzate = 0,
                Aggiunte = sourceRange.Rows,
                Finali = sourceRange.Rows,
            };
            infoRowsDestinazione.VerificaCoerenzaValori();
            Context.DebugInfoLogger.LogRigheSourceFiles(FileTypes.CN43N, infoRowsDestinazione);
            #endregion


            Context.CN43NFileEPPlusHelper.Close();
        }


        /// <summary>
        /// Aggiornamento qualitativo dei dati: per ogni riga del foglio sorgente, verifico se è già presente una riga corrispondente nel foglio destinazione (corrispondenza basata sul valore della colonna "WBSElement"). Se la riga è già presente, viene aggiornata con i dati della riga sorgente (copia e incolla del range), altrimenti viene aggiunta in coda alla destinazione. In questo modo vengono preservate tutte le righe esistenti che non trovano corrispondenza nel foglio sorgente e vengono aggiornate solo le righe che trovano corrispondenza.
        /// </summary>
        private void ImportDataFrom_CN43NFile_UpdateDuplicates()
        {
            // Foglio sorgente
            var sourceWorksheet = Context.CN43NFileEPPlusHelper.ExcelPackage.Workbook.Worksheets[1]; // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome

            // Foglio destinazione
            var destWorksheet = Context.DataSourceEPPlusHelper.ExcelPackage.Workbook.Worksheets[WorksheetNames.DATASOURCE_CN43N_DATA];
            var destHeadersRow = Context.Configurazione.DATASOURCE_CN43N_HEADERS_ROW;
            var numeroRigheIniziali = destWorksheet.Dimension.End.Row - destHeadersRow;

            // Prendo tutti gli elementi della colonna WBSElement dal foglio sorgente (sorted list (nome + indice riga)
            var sourceElements = Context.CN43NFileEPPlusHelper.GetValuesFromColumnsWithHeader(
                    worksheetName: sourceWorksheet.Name,
                    headersRow: Context.Configurazione.SOURCE_FILES_CN43N_HEADERS_ROW,
                    headerValue: Context.Configurazione.SOURCE_FILES_CN43N_HEADER_FOR_WBSLIST,
                    throwExceptionForMissingHeaders: true,
                    startHearderSearchFromColumn: 1);


            // Prendo tutti gli elementi della colonna WBSElement dal foglio destinazione (sorted list (nome + indice riga)
            var destElements = Context.DataSourceEPPlusHelper.GetValuesFromColumnsWithHeader(
                    worksheetName: WorksheetNames.DATASOURCE_CN43N_DATA,
                    headersRow: destHeadersRow,
                    headerValue: Context.Configurazione.DATASOURCE_FILES_CN43N_HEADER_FOR_WBSLIST,
                    throwExceptionForMissingHeaders: true,
                    startHearderSearchFromColumn: 1);


            #region Aggiornamento delle righe esistenti nella destinazione e conteggio delle righe riutilizzate
            // Per ogni elemento del foglio sorgente, verifico se è presente nel foglio destinazione:
            // se presente aggiorno la riga destinazione e cancello la riga sorgente
            int numeroRigheRiutilizzate = 0;
            foreach (var sourceItem in sourceElements)
            {
                var itemsAlreadyExistsInDest = destElements.Where(x => x.Valore == sourceItem.Valore).ToList();

                foreach (var itemAlreadyExistsInDest in itemsAlreadyExistsInDest)
                {
                    // Aggiorno la riga destinazione con i dati della riga sorgente (copia e incolla del range)
                    copyRowFromTo(sourceWorksheet, sourceItem.NumeroRiga, destWorksheet, itemAlreadyExistsInDest.NumeroRiga);

                    // Questa riga è giù stata usata, non la aggiungerò successivametne
                    sourceItem.Marked = true;
                    numeroRigheRiutilizzate++;
                }
            }
            #endregion


            #region Aggiunta delle righe non ancora presenti nella destinazione
            // Aggiungo in coda alla destinazione le eventuali righe rimanenti nel foglio sorgente (quelle che non hanno trovato corrispondenza nella destinazione)
            int righeAggiunte = 0;
            var righeDaAggiungere = sourceElements.Where(x => !x.Marked).ToList();
            foreach (var sourceItem in righeDaAggiungere)
            {
                // Copio la riga sorgente in coda alla destinazione
                copyRowFromTo(sourceWorksheet, sourceItem.NumeroRiga, destWorksheet, destWorksheet.Dimension.End.Row + 1);

                righeAggiunte++;
            }
            #endregion


            #region Log delle informazioni
            // Variabili per il conteggio delle righe elaborate
            var infoRowsDestinazione = new InfoRows
            {
                Iniziali = numeroRigheIniziali,
                Eliminate = 0,
                Preservate = 0,
                Riutilizzate = numeroRigheRiutilizzate,
                Aggiunte = righeAggiunte,
                Finali = destWorksheet.Dimension.End.Row - destHeadersRow,
            };
            infoRowsDestinazione.VerificaCoerenzaValori();
            Context.DebugInfoLogger.LogRigheSourceFiles(FileTypes.CN43N, infoRowsDestinazione);
            #endregion


            Context.CN43NFileEPPlusHelper.Close();
        }

        private void copyRowFromTo(ExcelWorksheet sourceWorksheet, int sourceRowIndex, ExcelWorksheet destWorksheet, int destRowIndex)
        {
            // Range sorgente
            var sourceRange = sourceWorksheet.Cells[
                            sourceRowIndex,                                                 // row start,
                            Context.Configurazione.SOURCE_FILES_CN43N_HEADERS_FIRST_COL,    // col start
                            sourceRowIndex,                              // row end
                            sourceWorksheet.Dimension.End.Column                            // col end
                            ];


            // Range destinazione
            var destRange = destWorksheet.Cells[
                            destRowIndex,                                         // row start,
                            Context.Configurazione.DATASOURCE_CN43N_HEADERS_FIRST_COL,  // col start
                            destRowIndex,                          // row end
                            sourceWorksheet.Dimension.End.Column                        // col end
                            ];

            // Incollo nel range destinazione
            destRange.Value = sourceRange.Value;
        }
    }
}