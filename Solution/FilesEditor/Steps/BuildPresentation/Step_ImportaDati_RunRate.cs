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
        internal override string StepName => "Step_ImportaDati_RunRate";

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


            #region Lettura del valore per la variabile "cell_valore_RR"
            var valore_cell_perc_RR = (double?) Context.DataSourceEPPlusHelper.GetValue(WorksheetNames.DATASOURCE_RUN_RATE_DATA, Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_ROW + 1, Context.PeriodMont);
            if (valore_cell_perc_RR.HasValue)
            {
                const string VARIABLE_NAME_CELL_PERC_RR = "cell_perc_RR";
                Context.DataSourceEPPlusHelper.SetVariableInNameManager(VARIABLE_NAME_CELL_PERC_RR, valore_cell_perc_RR);
            }
            #endregion


            Context.RunRateFileEPPlusHelper.Close();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }
    }
}