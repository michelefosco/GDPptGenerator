using ReportRefresher.Entities;
using System.Linq;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Inserimento di eventuali nuovi fornitori forniti dall'interfaccia
    /// </summary>
    internal class Step_Inserimento_NuoviFornitori_DaInterfaccia : Step_Inserimento_NuoviFornitori_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            if (context.UpdateReportsInput.FornitoriDaAggiungere != null && context.UpdateReportsInput.FornitoriDaAggiungere.Any())
            {
                InserimentoNuoviFornitori(context.InfoFileReport, context.Configurazione, context.FornitoriCensitiInReport, context.RepartiCensitiInReport, context.CategorieFornitori, context.UpdateReportsInput.FornitoriDaAggiungere);
                context.DebugInfoLogger.LogFornitoriAggiuntiDaInterfaccia(context.UpdateReportsInput.FornitoriDaAggiungere, context.RepartiCensitiInReport);
                context.DebugInfoLogger.LogText("Aggiunti i nuovi fornitori ricevuti dall'interfaccia utente", context.UpdateReportsInput.FornitoriDaAggiungere.Count);
            }

            return null;
        }
    }
}