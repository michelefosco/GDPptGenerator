using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Enums;
using System;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_ImportaDati_RunRate : StepBase
    {
        public override string StepName => "Step_ImportaDati_RunRate";

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

        public Step_ImportaDati_RunRate(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            // Foglio sorgente
            // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome
            var sourceWorksheet = Context.RunRateFileEPPlusHelper.ExcelPackage.Workbook.Worksheets[1];

            // Foglio destinazione
            var destWorksheet = Context.DataSourceEPPlusHelper.ExcelPackage.Workbook.Worksheets[WorksheetNames.DATASOURCE_RUN_RATE_DATA];

            // Range sorgente
            var sourceRange = sourceWorksheet.Cells[
                    Context.Configurazione.SOURCE_FILES_RUNRATE_HEADERS_ROW + 1,  // row start
                    Context.Configurazione.SOURCE_FILES_RUNRATE_HEADERS_FIRST_COL, // col start 
                    Context.Configurazione.SOURCE_FILES_RUNRATE_HEADERS_ROW + 1,   // row end
                    Context.Configurazione.SOURCE_FILES_RUNRATE_HEADERS_FIRST_COL + 11    // col end
                    ];

            // Incollo nel range destinazione
            var destRange = destWorksheet.Cells[
                    Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_ROW + 1,  // row start
                    Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_FIRST_COL, // col start 
                    Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_ROW + 1,   // row end
                    Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_FIRST_COL + 11    // col end
                    ];
            destRange.Value = sourceRange.Value;

            destWorksheet.Select(destWorksheet.Cells[1, 1]);
            Context.RunRateFileEPPlusHelper.Close();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }
    }
}