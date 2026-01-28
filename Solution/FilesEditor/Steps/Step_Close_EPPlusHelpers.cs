using FilesEditor.Entities;
using FilesEditor.Enums;


namespace FilesEditor.Steps
{
    /// <summary>
    /// Step che se raggiunto indica l'esito Success dell'intero processo
    /// </summary>
    internal class Step_Close_EPPlusHelpers : StepBase
    {
        internal override string StepName => "Step_Close_EPPlusHelpers";
        public Step_Close_EPPlusHelpers(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            // Chiudo tutti gli helper EPPlus aperti
            // Datasouce file
            Context.DataSourceEPPlusHelper.Close();

            // Input source files
            if (!string.IsNullOrEmpty(Context.FileBudgetPath))
            { Context.BudgetFileEPPlusHelper.Close(); }

            if (!string.IsNullOrEmpty(Context.FileForecastPath))
            { Context.ForecastFileEPPlusHelper.Close(); }

            if (!string.IsNullOrEmpty(Context.FileRunRatePath))
            { Context.RunRateFileEPPlusHelper.Close(); }

            if (!string.IsNullOrEmpty(Context.FileSuperDettagliPath))
            { Context.SuperdettagliFileEPPlusHelper.Close(); }

            if (!string.IsNullOrEmpty(Context.FileCN43NPath))
            { Context.CN43NFileEPPlusHelper.Close(); }

            return EsitiFinali.Undefined;
        }
    }
}