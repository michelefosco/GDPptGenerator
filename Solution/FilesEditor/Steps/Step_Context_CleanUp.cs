using FilesEditor.Entities;
using FilesEditor.Enums;


namespace FilesEditor.Steps
{
    /// <summary>
    /// Step che se raggiunto indica l'esito Success dell'intero processo
    /// </summary>
    internal class Step_Context_CleanUp : StepBase
    {
        internal override string StepName => "Step_Context_CleanUp";
        public Step_Context_CleanUp(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            Context.DataSourceEPPlusHelper.Close();
            Context.BudgetFileEPPlusHelper.Close();
            Context.ForecastFileEPPlusHelper.Close();
            Context.RunRateFileEPPlusHelper.Close();
            Context.SuperdettagliFileEPPlusHelper.Close();

            // Essendo un file opzionale, verifico che l'helper non sia null prima di chiuderlo
            if (!string.IsNullOrEmpty(Context.FileCN43NPath))
            { Context.CN43NFileEPPlusHelper.Close(); }

            Serilog.Log.Information("CloseAndFlush...");
            Serilog.Log.CloseAndFlush();

            return EsitiFinali.Undefined;
        }
    }
}