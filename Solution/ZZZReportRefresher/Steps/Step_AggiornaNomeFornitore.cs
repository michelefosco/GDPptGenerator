using ReportRefresher.Entities;
using ReportRefresher.Entities.Exceptions;
using ReportRefresher.Enums;
using System;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Lettura categorie fornitori
    /// </summary>
    internal class Step_AggiornaNomeFornitore : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            var siglaFornitore = context.Parameters["siglaFornitore"].ToString();
            var nuovoNomeFornitore = context.Parameters["nuovoNomeFornitore"].ToString();

            UpdateFornitore(context.InfoFileReport, context.Configurazione, siglaFornitore, nuovoNomeFornitore);
            return null;
        }

        private void UpdateFornitore(InfoFileReport infoFileReport, Configurazione configurazione, string siglaFornitore, string nuovoNomeFornitore)
        {
            var worksheetName = infoFileReport.WorksheetName_AnagraficaFornitori; // Anagrafica fornitori

            var ultimaRigaUsataNelFoglio = infoFileReport.EPPlusHelper.GetRowsLimit(worksheetName);
            var rigaCorrente = configurazione.AnagraficaFornitori_PrimaRigaFornitori;
            while (rigaCorrente <= ultimaRigaUsataNelFoglio)
            {
                var siglaFornitoreRigaCorrente = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.AnagraficaFornitori_ColonnaSigle);

                if (siglaFornitore.Equals(siglaFornitoreRigaCorrente, StringComparison.InvariantCultureIgnoreCase))
                {
                    // aggiorno la colonna dei nomi fornitori con il nuovo valore
                    infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaCorrente, configurazione.AnagraficaFornitori_ColonnaNomi, nuovoNomeFornitore);
                    return;
                }
                // vado alla riga successiva
                rigaCorrente++;
            }

            // se sono arrivato fin qui vuol dire che la sigla del fonritore cercato non è stata trovata
            throw new ManagedException(
                tipologiaErrore: TipologiaErrori.DatoMancante,
                tipologiaCartella: TipologiaCartelle.ReportInput,
                nomeDatoErrore: NomiDatoErrore.SiglaFornitore,
                worksheetName: worksheetName,
                rigaCella: rigaCorrente,
                colonnaCella: configurazione.AnagraficaFornitori_ColonnaSigle,
                dato: siglaFornitore
                //percorsoFile: null
                );
        }
    }
}