using FilesEditor.Entities;
using FilesEditor.Enums;
using System;

namespace FilesEditor.Steps
{
    /// <summary>
    /// Step che se raggiunto indica l'esito Success dell'intero processo
    /// </summary>
    internal class Step_Context_CleanUp: StepBase
    {
        public override string StepName => "Step_Context_CleanUp";

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

        public Step_Context_CleanUp(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            Context.DataSourceEPPlusHelper.Close();
            Context.BudgetFileEPPlusHelper.Close();
            Context.ForecastFileEPPlusHelper.Close();
            Context.RunRateFileEPPlusHelper.Close();
            Context.SuperdettagliFileEPPlusHelper.Close();
            Context.CN43NFileEPPlusHelper.Close();

            return EsitiFinali.Undefined;
        }
    }
}