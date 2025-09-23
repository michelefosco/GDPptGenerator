using ReportRefresher.Entities;
using System.Linq;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Inserimento degli eventuali fornitori Presenti solo in "Lista dati" per i quali però sono state trovate spese nel file controller.
    /// Per questi fornitori si può non passare dall'interfaccia
    /// </summary>
    internal class Step_Inserimento_NuoviFornitori_PresentiSoloInListaDati : Step_Inserimento_NuoviFornitori_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            // I fornitori censiti, inizialmente marcati come "PresenteSoloInListaDati" e successivamente
            // marcati anche come "DeveEsserePresenteNeiReport" in quanto per essi sono state individuate delle spese
            // devono essere aggiunti nelle varie posisioni dei report, esattamente come se si trattasse di nuovi fornitori
            var fornitoriDaAggiungere = context.FornitoriCensitiInReport.Where(_ => _.PresenteSoloInListaDati && _.DeveEsserePresenteNeiReport).ToList();

            //    if (context.FornitoriCensitiInReport.Any(_ => _.PresenteSoloInListaDati && _.DeveEsserePresenteNeiReport))
            if (fornitoriDaAggiungere.Any())
            {
                InserimentoNuoviFornitori(context.InfoFileReport, context.Configurazione, context.FornitoriCensitiInReport, context.RepartiCensitiInReport, context.CategorieFornitori, fornitoriDaAggiungere);
                context.DebugInfoLogger.LogFornitoriAggiuntiDaListaDati(fornitoriDaAggiungere, context.RepartiCensitiInReport);
                context.DebugInfoLogger.LogText("Aggiunti nuovi fornitori trovati in Lista Dati", fornitoriDaAggiungere.Count);
            }

            return null;
        }
    }
}