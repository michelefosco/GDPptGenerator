using ReportRefresher.Constants;
using ReportRefresher.Entities;
using ReportRefresher.Entities.Exceptions;
using ReportRefresher.Enums;
using System.Collections.Generic;
using System;
using System.Linq;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Lettura info fornitori censiti
    /// </summary>
    internal class Step_Lettura_CostiOrari_CategorieDeiFornitori : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            // Lettura dei costi orari dal foglio "Lista Dati" del file "Report"
            GetCostiOrariAndCategorieDeiFornitori(context.InfoFileReport, context.Configurazione, context.FornitoriCensitiInReport, context.RepartiCensitiInReport, context.CategorieFornitori);
            context.UpdateReportsOutput.SettaFornitoriCensiti(context.FornitoriCensitiInReport);

            context.DebugInfoLogger.LogText("Lettura fornitori censiti", context.FornitoriCensitiInReport.Count);
            context.DebugInfoLogger.LogText("Lettura fornitori censiti (Presenti solo in Lista dati)", context.FornitoriCensitiInReport.Count(_ => _.PresenteSoloInListaDati));
            context.DebugInfoLogger.LogFornitoriCensiti(context.FornitoriCensitiInReport, context.RepartiCensitiInReport);

            return null;
        }
        private void GetCostiOrariAndCategorieDeiFornitori(InfoFileReport infoFileReport, Configurazione configurazione, List<FornitoreCensito> fornitoriCensitiInReport, List<Reparto> repartiCensitiInReport, List<string> categorieFornitori/*, out List<FornitoreCensitoInReport> fornitoriPresentiInListaDati*/)
        {
            //fornitoriPresentiInListaDati = new List<FornitoreCensitoInReport>();
            var worksheetName = infoFileReport.WorksheetName_ListaDati; // Lista dati

            #region lettura dei reparti dalla tabella dei costi ad ok per reparto
            var rigaReparti = configurazione.ListaDati_RigaReparti;
            var nomiRepartiSuReport = new List<string>();
            var colonnaCorrente = configurazione.ListaDati_ColonnaReparto;
            while (true)
            {
                var nomeReparto = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaReparti, colonnaCorrente);
                if (string.IsNullOrWhiteSpace(nomeReparto)) // mi interrompo con l'ultima cella con eventuale valore blank
                { break; }

                // Il reparto deve essere censito sul file Report altrimenti rappresenta un problema
                var repartoCensitoInReport = repartiCensitiInReport.SingleOrDefault(_ => _.Nome.Equals(nomeReparto, StringComparison.InvariantCultureIgnoreCase));
                if (repartoCensitoInReport == null)
                {
                    // FRANCESCO dice: In Lista Dati ci sono meno reparti poiché ai fini dei report non occorre il censimento di tutti i reparti di Ufficio Tecnico. Nelle voci di spesa (ovvero fogli ‘act_solo_cdc’, ‘FBL3N’, ‘committment_solo_cdc’, ‘ME5A’, ‘ME2N’) ci sono solo alcuni dei reparti. Il codice potrebbe censirli tutti, e poi usarli come Dizionario, prendendo solo quelli che appaiono nelle voci di spesa.
                    // Bisogna segnalare all’utente solo se tra le voci di spesa del file dei controller(non nel foglio RECAP) appare un centro di costo afferente ad un reparto NON presente in lista dati. In quel caso toccherà a noi aggiungerlo nelle tabelle dei reports.
                    throw new ManagedException(
                        TipologiaErrori.DatoNonValido,
                        TipologiaCartelle.ReportInput,
                        $"Il foglio '{worksheetName}' ha un reparto non correttamente censito sul file controller ('{nomeReparto}')",
                        worksheetName,
                        null,
                        rigaReparti,
                        colonnaCorrente,
                        nomeReparto
                        );
                }

                nomiRepartiSuReport.Add(repartoCensitoInReport.Nome);
                colonnaCorrente++;
            }
            #endregion

            var ultimaRigaUsataNelFoglio = infoFileReport.EPPlusHelper.GetRowsLimit(worksheetName);
            var rigaCorrente = configurazione.ListaDati_PrimaRigaFornitori - 1;
            while (rigaCorrente <= ultimaRigaUsataNelFoglio)
            {
                // vado alla riga successiva
                rigaCorrente++;

                var siglaFornitore = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.ListaDati_ColonnaSiglaFornitoriTabellaSX);
                if (siglaFornitore == null)
                { continue; }

                #region Valuto la presenza di sigle fornitori presenti solo in "Lista dati"
                var fornitoreCensito = fornitoriCensitiInReport.SingleOrDefault(_ => _.SiglaInReport.Equals(siglaFornitore, StringComparison.InvariantCultureIgnoreCase));
                if (fornitoreCensito == null)
                {
                    // caso di una sigla di fornitore presente in "Lista dati" ma non censito in "Anagrafica fornitore".
                    // Francesco dice: questo caso deve essere permesso in quanto utile per la pianificazione di nuovi fornitori.
                    // aggiungo il fornitore alla lista dei fornitori censiti ma marcato come "PresenteSoloInListaDati"
                    fornitoreCensito = new FornitoreCensito(siglaInReport: siglaFornitore, nomeSuController: siglaFornitore, presenteSoloInListaDati: true);
                    fornitoriCensitiInReport.Add(fornitoreCensito);
                }
                #endregion

                var objectCostoOrarioDefault = infoFileReport.EPPlusHelper.GetValue(worksheetName, rigaCorrente, configurazione.ListaDati_ColonnaCostiOrariStandard) as double?;
                if (!objectCostoOrarioDefault.HasValue)
                {
                    // caso in cui risulti mancante il prezzo di default per un fornitore
                    throw new ManagedException(
                        tipologiaErrore: TipologiaErrori.DatoMancante,
                        tipologiaCartella: TipologiaCartelle.ReportInput,
                        messaggioPerUtente: string.Format(MessaggiErrorePerUtente.CorsoOrarioFornitoreNonValido, siglaFornitore),
                        worksheetName: worksheetName,
                        rigaCella: rigaCorrente,
                        colonnaCella: configurazione.ListaDati_ColonnaCostiOrariStandard,
                        nomeDatoErrore: NomiDatoErrore.CostoOrarioFornitore,
                        dato: null,
                        percorsoFile: null
                             );
                }

                // Non posso avere costi negativi. Francesco dice che il costo orario di default può essere a zero (di solito usato per i fornitori Lump sum)
                if (objectCostoOrarioDefault.Value < 0)
                {
                    // caso in cui risulti mancante il presso di default per un fornitore
                    throw new ManagedException(
                        tipologiaErrore: TipologiaErrori.DatoNonValido,
                        tipologiaCartella: TipologiaCartelle.ReportInput,
                        messaggioPerUtente: string.Format(MessaggiErrorePerUtente.CorsoOrarioFornitoreNonValido, siglaFornitore),
                        worksheetName: worksheetName,
                        rigaCella: rigaCorrente,
                        colonnaCella: configurazione.ListaDati_ColonnaCostiOrariStandard,
                        nomeDatoErrore: NomiDatoErrore.CostoOrarioFornitore,
                        dato: objectCostoOrarioDefault.Value.ToString(),
                        percorsoFile: null
                           );
                }

                #region Assegno la categoria al fornitore
                var categoriaFornitori = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.ListaDati_ColonnaCategoriaFornitori);
                if (string.IsNullOrWhiteSpace(categoriaFornitori))
                {
                    //situazione in cui in "Lista dati" non è indicata la categoria del fornitore
                    throw new ManagedException(
                        tipologiaErrore: TipologiaErrori.DatoMancante,
                        tipologiaCartella: TipologiaCartelle.ReportInput,
                        worksheetName: worksheetName,
                        rigaCella: rigaCorrente,
                        colonnaCella: configurazione.ListaDati_ColonnaCategoriaFornitori,
                        nomeDatoErrore: NomiDatoErrore.CategoriaFornitore,
                        percorsoFile: null
                        );
                }

                if (categorieFornitori.SingleOrDefault(_ => _.Equals(categoriaFornitori, StringComparison.InvariantCultureIgnoreCase)) == null)
                {
                    // situazione in cui in "Lista dati" è indicata la categoria del fornitore non censita sul foglio "Reportistica per tipologia"
                    throw new ManagedException(
                        tipologiaErrore: TipologiaErrori.DatoNonValido,
                        tipologiaCartella: TipologiaCartelle.ReportInput,
                        worksheetName: worksheetName,
                        rigaCella: rigaCorrente,
                        colonnaCella: configurazione.ListaDati_ColonnaCategoriaFornitori,
                        nomeDatoErrore: NomiDatoErrore.CategoriaFornitore,
                        dato: categoriaFornitori,
                        percorsoFile: null
                        );
                }
                fornitoreCensito.SettaCategoria(categoriaFornitori);
                #endregion

                #region Assegno i costi orari al fornitore
                var prezziSpecificiPerReparto = new Dictionary<string, double>();
                for (var r = 0; r < nomiRepartiSuReport.Count; r++)
                {
                    var valueCostoOrarioPerReparto = infoFileReport.EPPlusHelper.GetValue(worksheetName, rigaCorrente, configurazione.ListaDati_PrimaColonnaReparti + r) as double?;
                    if (valueCostoOrarioPerReparto.HasValue)
                    {
                        prezziSpecificiPerReparto.Add(nomiRepartiSuReport[r], valueCostoOrarioPerReparto.Value);
                    }
                }
                fornitoreCensito.SettaCostiOrari(objectCostoOrarioDefault.Value, prezziSpecificiPerReparto);
                #endregion
            }
        }
    }
}