using ReportRefresher.Entities;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Attivazione dell'opzione "refreshOnLoad" su tutte le Pivot Tables
    /// </summary>
    internal class Step_AttivazioneOpzioneRefreshOnLoad : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            if (context.Configurazione.AttivazioneOpzione_RefreshOnLoad_SuTutteLePivotTable)
            {
                var numeroDiPivotTableAggiornate = context.InfoFileReport.EPPlusHelper.SetRefreshOnLoadForAllPivotTables();
                context.DebugInfoLogger.LogText("Attivazione opzione 'refreshOnLoad' delle pivot table", numeroDiPivotTableAggiornate);
            }

            return null;
        }
    }
}