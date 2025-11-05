using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Enums;
using OfficeOpenXml;
using System;
using System.IO;

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
            var packageSource = new ExcelPackage(new FileInfo(Context.FileRunRatePath));
            var sourceWorksheet = packageSource.Workbook.Worksheets[WorksheetNames.SOURCEFILE_RUN_RATE_DATA];

            // Foglio destinazione
            var destWorksheet = Context.EpplusHelperDataSource.ExcelPackage.Workbook.Worksheets[WorksheetNames.DATASOURCE_RUN_RATE_DATA];

            // Range sorgente
            var sourceRange = sourceWorksheet.Cells[
                    Context.Configurazione.SOURCE_FILES_RUNRATE_HEADERS_ROW + 1,  // row start
                    Context.Configurazione.SOURCE_FILES_RUNRATE_HEADERS_FIRST_COL, // col start 
                    Context.Configurazione.SOURCE_FILES_RUNRATE_HEADERS_ROW + 1,   // row end
                    Context.Configurazione.SOURCE_FILES_RUNRATE_HEADERS_FIRST_COL + 11    // col end
                    ];

            // Incollo nel range destinazione
            destWorksheet.Cells[
                    Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_ROW + 1,  // row start
                    Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_FIRST_COL, // col start 
                    Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_ROW + 1,   // row end
                    Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_FIRST_COL + 11    // col end
                    ].Value = sourceRange.Value;

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }
   }
}