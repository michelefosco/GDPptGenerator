using FilesEditor.Entities;
using FilesEditor.Enums;
using System;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_SalvaFile_DataSource: StepBase
    {
        public override string StepName => "Step_SalvaFile_DataSource";

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

        public Step_SalvaFile_DataSource(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoStepTask()
        {
            Context.EpplusHelperDataSource.Save();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }
    }
}