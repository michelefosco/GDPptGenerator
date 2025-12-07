using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Enums;
using System;


namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_ImportaDati_CN43N : StepBase
    {
        public override string StepName => "Step_ImportaDati_CN43N";

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

            // Variabili per il conteggio delle righe elaborate
            var infoRowsDestinazione = new InfoRows();
            infoRowsDestinazione.Iniziali = destWorksheet.Dimension.End.Row - destHeadersRow;

            // Elimino le righe esistenti (tranne l'intestazione)
            var numeroRigheDaEliminare = destWorksheet.Dimension.End.Row - destHeadersRow;
            destWorksheet.DeleteRow(destHeadersRow + 1, numeroRigheDaEliminare, true);
            infoRowsDestinazione.Eliminate = numeroRigheDaEliminare;

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


            // Log delle informazioni
            infoRowsDestinazione.Aggiunte = sourceRange.Rows;
            infoRowsDestinazione.Preservate = 0;
            infoRowsDestinazione.Riutilizzate = 0;
            infoRowsDestinazione.Finali = destWorksheet.Dimension.End.Row - destHeadersRow;
            infoRowsDestinazione.VerificaCoerenzaValori();
            Context.DebugInfoLogger.LogRigheSourceFiles(FileTypes.CN43N, infoRowsDestinazione);

            destWorksheet.Select(destWorksheet.Cells[1, 1]);
            Context.CN43NFileEPPlusHelper.Close();
        }
    }
}