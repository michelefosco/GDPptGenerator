using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Enums;
using OfficeOpenXml;
using System;

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


        private void ImportDataFrom_CN43NFile_UpdateDuplicates()
        {

            //todo: tenere traccia delle righe modificate

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
            var destElements = Context.CN43NFileEPPlusHelper.GetValuesFromColumnsWithHeader(
                    worksheetName: WorksheetNames.DATASOURCE_CN43N_DATA,
                    headersRow: destHeadersRow,
                    headerValue: Context.Configurazione.DATASOURCE_FILES_CN43N_HEADER_FOR_WBSLIST,
                    throwExceptionForMissingHeaders: true,
                    startHearderSearchFromColumn: 1);


            int numeroRigheRiutilizzate = 0;

            // Per ogni elemento del foglio sorgente, verifico se è presente nel foglio destinazione:
            // se presente aggiorno la riga destinazione e cancello la riga sorgente
            foreach (var sourceItem in sourceElements)
            {
                if(destElements.ContainsKey(sourceItem.Key))
                {
                    // Aggiorno la riga destinazione con i dati della riga sorgente (copia e incolla del range)
                    copyRowFromTo(sourceWorksheet, sourceItem.Value, destWorksheet, destElements[sourceItem.Key]);

                    numeroRigheRiutilizzate++;

                    // Cancello la riga sorgente appena copiata, in modo da avere alla fine solo le righe che non hanno trovato corrispondenza nella destinazione (quelle che andranno aggiunte in coda alla fine del processo)
                    sourceElements.Remove(sourceItem.Key);                   
                }                
            }

            // Aggiungo in coda alla destinazione le eventuali righe rimanenti nel foglio sorgente (quelle che non hanno trovato corrispondenza nella destinazione)
            int righeAggiunte = 0;
            //todo: aggiungere righe in coda alla destinazione, invece che incollarle tutte insieme (come fatto ora) per tenere traccia del numero di righe aggiunte

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