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
            Context.BudgetFileEPPlusHelper.Close();
            Context.ForecastFileEPPlusHelper.Close();
            Context.RunRateFileEPPlusHelper.Close();
            Context.SuperdettagliFileEPPlusHelper.Close();
            // Essendo un file opzionale, verifico che l'helper non sia null prima di chiuderlo
            if (!string.IsNullOrEmpty(Context.FileCN43NPath))
            { Context.CN43NFileEPPlusHelper.Close(); }

            return EsitiFinali.Undefined;
        }
    }
}