using ReportRefresher.Entities;
using ReportRefresher.Enums;
using System.Linq;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Valutazione presenza di "Nuovi fornitori" con dati mancanti (non recuperabili da "Lista dati")
    /// </summary>
    internal class Step_ValutazionePresenzaNuoviFornitori: Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            if (context.FornitoriNonCensitiInReport != null && context.FornitoriNonCensitiInReport.Any())
            {
                context.UpdateReportsOutput.SettaFornitoriNonCensitiInReport(context.FornitoriNonCensitiInReport);
                context.DebugInfoLogger.LogFornitoriNonCensitiInReport(context.FornitoriNonCensitiInReport);
                context.DebugInfoLogger.LogText("Elaborazione interrotta a causa di fornitori con dati mancanti", context.FornitoriNonCensitiInReport.Count);

                // Se ho dei fornitori da aggiungere, mi fermo qui per chiedere le info aggiuntive all'interfaccia
                context.UpdateReportsOutput.SettaEsitoFinale(EsitiFinali.DatiFornitoriMancanti);

                ChiusuraFileDebug(context);
                return context.UpdateReportsOutput;
            }

            return null;
        }
    }
}