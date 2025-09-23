using ReportRefresher.Entities;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Gestione tabella "Previsione a finire"
    /// </summary>
    internal class Step_TabellaPrevisioneAfinire : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            AggiornaFormuleSuTabellaPrevisioneAfinire(context.InfoFileReport, context.Configurazione, context.UpdateReportsInput.Periodo);
            context.DebugInfoLogger.LogText("Aggiornamento delle formule sul foglio Previsione a finire", context.UpdateReportsInput.Periodo);

            return null;
        }

        private void AggiornaFormuleSuTabellaPrevisioneAfinire(InfoFileReport infoFileReport, Configurazione configurazione, int periodo)
        {
            var worksheetName = infoFileReport.WorksheetName_PrevisioneAfinire;  // Previsione a finire

            #region Aggiornamento delle formule nella prima cella della colonna
            // Francesco dice: 
            // Praticamente bisogna sostituire solo un fattore a seconda del mese di lancio:
            // =IF($G6<>"non-visibile"; K6 * 12 / 9; "non-visibile")
            // Corrisponde sempre al numero del mese precedente.
            // Esempio: a Ottobre è 9; a Dicembre è 11; a Gennaio è 12 ecc.

            // La formula: =SE($G6<>"non-visibile";K6*12/<X>;"non-visibile")
            // va aggiornata quindi sostituendo <X> con il giusto numero

            var xNellaFormula = (periodo == 1) ? 12 : periodo - 1;

            // Aggiorno la formula nella prima cella della colonna
            var formulaDaIncollare = string.Format(configurazione.PrevisioneAfinire_FormulaRanRateOre, xNellaFormula);
            infoFileReport.EPPlusHelper.SetFormula(worksheetName, configurazione.PrevisioneAfinire_Riga_PrimaConDati, configurazione.PrevisioneAfinire_ColonnaRanRateOre, formulaDaIncollare);
            #endregion

            #region  Copia e incolla delle formule dalla prima riga a tutte le altre nella colonna
            // tolgo 1 alla prima riga vuota incontrata per tornare all'ultima riga con valori al suo interno
            var primaRigaInCuiCopiareLeFormule = configurazione.PrevisioneAfinire_Riga_PrimaConDati + 1;
            var ultimaRigaInCuiCopiareLeFormule = infoFileReport.EPPlusHelper.GetLastUsedRowForColumn(worksheetName, primaRigaInCuiCopiareLeFormule, configurazione.PrevisioneAfinire_ColonnaRanRateOre);

            for (var rigaDaAggiornare = primaRigaInCuiCopiareLeFormule; rigaDaAggiornare <= ultimaRigaInCuiCopiareLeFormule; rigaDaAggiornare++)
            {
                infoFileReport.EPPlusHelper.CopiaFormula(worksheetName, configurazione.PrevisioneAfinire_Riga_PrimaConDati, configurazione.PrevisioneAfinire_ColonnaRanRateOre,
                                                                     rigaDaAggiornare, configurazione.PrevisioneAfinire_ColonnaRanRateOre);
            }
            #endregion
        }
    }
}