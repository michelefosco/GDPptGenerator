using FilesEditor.Entities;
using FilesEditor.Enums;
using System;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// Attivazione dell'opzione "refreshOnLoad" su tutte le Pivot Tables
    /// </summary>
    internal class Step_AttivazioneOpzioneRefreshOnLoad : StepBase
    {
        public override string StepName => "Step_AttivazioneOpzioneRefreshOnLoad";

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

        public Step_AttivazioneOpzioneRefreshOnLoad(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            //todo: verificarne il funzionamento quando ci saranno pivet tables nel datasource file
            var numeroDiPivotTableAggiornate = Context.EpplusHelperDataSource.SetRefreshOnLoadForAllPivotTables();
            Context.DebugInfoLogger.LogText("Attivazione opzione 'refreshOnLoad' delle pivot table", numeroDiPivotTableAggiornate);

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }
    }
}