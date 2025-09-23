using ReportRefresher.Entities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Lettura categorie fornitori
    /// </summary>
    internal class Step_Lettura_CategorieFornitori : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            context.CategorieFornitori = GetListaCategorieFornitori(context.InfoFileReport, context.Configurazione);
            context.UpdateReportsOutput.SettaCategorieFornitori(context.CategorieFornitori);
            context.DebugInfoLogger.LogCategorieFornitori(context.CategorieFornitori);

            return null;
        }

        private List<string> GetListaCategorieFornitori(InfoFileReport infoFileReport, Configurazione configurazione)
        {
            var worksheetName = infoFileReport.WorksheetName_ReportisticaPerTipologia; // Reportistica per tipologia

            var ultimaRigaUsataNelFoglio = infoFileReport.EPPlusHelper.GetRowsLimit(worksheetName);
            var rigaCorrente = configurazione.ReportisticaPerTipologia_PrimaRigaCategorieFornitori;

            var categorieFornitore = new List<string>();
            while (rigaCorrente <= ultimaRigaUsataNelFoglio)
            {
                var categoriaForntiore = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.ReportisticaPerTipologia_ColonnaCategorieFornitori);

                // mi fermo quando incontro una cella con valore nullo oppure contenente il testo "Totale"
                if (categoriaForntiore == null || categoriaForntiore.Equals("Totale", StringComparison.CurrentCultureIgnoreCase))
                { break; }

                categorieFornitore.Add(categoriaForntiore);

                // vado alla riga successiva
                rigaCorrente++;
            }

            // tolto i duplicati ed ordino alfabeticamente la lista
            return categorieFornitore.Distinct().OrderBy(_ => _).ToList();
        }
    }
}