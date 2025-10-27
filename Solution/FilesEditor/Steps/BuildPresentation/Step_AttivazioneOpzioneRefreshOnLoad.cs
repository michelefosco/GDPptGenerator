using FilesEditor.Entities;
using FilesEditor.Enums;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// Attivazione dell'opzione "refreshOnLoad" su tutte le Pivot Tables
    /// </summary>
    internal class Step_AttivazioneOpzioneRefreshOnLoad : StepBase
    {
        public Step_AttivazioneOpzioneRefreshOnLoad(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            //todo: verificare uk fynzionamento
            var numeroDiPivotTableAggiornate = Context.EpplusHelperDataSource.SetRefreshOnLoadForAllPivotTables();
            Context.DebugInfoLogger.LogText("Attivazione opzione 'refreshOnLoad' delle pivot table", numeroDiPivotTableAggiornate);

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }
    }
}