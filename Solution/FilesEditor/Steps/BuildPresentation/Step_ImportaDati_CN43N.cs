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
            if (string.IsNullOrEmpty(Context.FileCN43NPath))
            {
                Context.DebugInfoLogger.LogWarning($"File CN43N non specificato. Si salta l'importazione dei dati CN43N.");
            }
            else
            {
                ImportDataFrom_CN43NFile();
                Context.CN43NFileEPPlusHelper.Close();
            }

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void ImportDataFrom_CN43NFile()
        {
            // Foglio sorgente
            // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome
            var sourceWorksheet = Context.CN43NFileEPPlusHelper.ExcelPackage.Workbook.Worksheets[1];

            // Foglio destinazione
            var destWorksheet = Context.DataSourceEPPlusHelper.ExcelPackage.Workbook.Worksheets[WorksheetNames.DATASOURCE_CN43N_DATA];

            // Elimino le righe esistenti (tranne l'intestazione)
            destWorksheet.DeleteRow(Context.Configurazione.DATASOURCE_FILES_CN43N_HEADERS_ROW + 1, destWorksheet.Dimension.End.Row - Context.Configurazione.DATASOURCE_FILES_CN43N_HEADERS_ROW, true);

            // Range sorgente
            var sourceRange = sourceWorksheet.Cells[
                           Context.Configurazione.SOURCE_FILES_CN43N_HEADERS_ROW,       // row start,
                           Context.Configurazione.SOURCE_FILES_CN43N_HEADERS_FIRST_COL, // col start
                            sourceWorksheet.Dimension.End.Row,      // row end
                            sourceWorksheet.Dimension.End.Column    // col end
                            ];

            // Incollo nel range destinazione
            destWorksheet.Cells[
                            Context.Configurazione.DATASOURCE_FILES_CN43N_HEADERS_ROW,          // row start,
                            Context.Configurazione.DATASOURCE_FILES_CN43N_HEADERS_FIRST_COL,    // col start
                            sourceWorksheet.Dimension.End.Row,      // row end
                            sourceWorksheet.Dimension.End.Column    // col end
                    ].Value = sourceRange.Value;
        }
    }
}