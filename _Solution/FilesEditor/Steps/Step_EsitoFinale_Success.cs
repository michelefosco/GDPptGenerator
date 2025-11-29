using FilesEditor.Entities;
using FilesEditor.Enums;
using System;

namespace FilesEditor.Steps
{
    /// <summary>
    /// Step che se raggiunto indica l'esito Success dell'intero processo
    /// </summary>
    internal class Step_EsitoFinale_Success : StepBase
    {
        public override string StepName => "Step_EsitoFinale_Success";

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

        public Step_EsitoFinale_Success(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            return EsitiFinali.Success;
        }
    }
}