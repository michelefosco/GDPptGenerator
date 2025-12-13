using FilesEditor.Entities;
using FilesEditor.Enums;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// Attivazione dell'opzione "refreshOnLoad" su tutte le Pivot Tables
    /// </summary>
    internal class Step_AttivazioneOpzioneRefreshOnLoad : StepBase
    {
        internal override string StepName => "Step_AttivazioneOpzioneRefreshOnLoad";

        public Step_AttivazioneOpzioneRefreshOnLoad(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            //todo: verificarne il funzionamento quando ci saranno pivot tables nel datasource file
            var numeroDiPivotTableAggiornate = Context.DataSourceEPPlusHelper.SetRefreshOnLoadForAllPivotTables();
            Context.DebugInfoLogger.LogText("Attivazione opzione 'refreshOnLoad' delle pivot table", numeroDiPivotTableAggiornate);

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }
    }
}