using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Enums;

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
                ImportDataFrom_CN43NFile();
            }

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void ImportDataFrom_CN43NFile()
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
    }
}