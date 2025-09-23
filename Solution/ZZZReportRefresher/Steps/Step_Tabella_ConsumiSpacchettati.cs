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
    /// Gestione tabella "Consumi Spacchettati"
    /// </summary>
    internal class Step_Tabella_ConsumiSpacchettati : Step_Base
    {
        private void tmpCheckVersioneConModificaInFoglioConsumiSpacchettatiCancellareAllaProssimaRealese(StepContext context)
        {
            var worksheetName = context.InfoFileReport.WorksheetName_ConsumiSpacchettati; // "Consumi Spacchettati"
            int riga = 16;
            int colonna1 = context.Configurazione.ConsumiSpacchettati_ColonnaReparti;
            int colonna2 = 20; //T

            var value1 = context.InfoFileReport.EPPlusHelper.GetValue(worksheetName, riga, colonna1);
            var value2 = context.InfoFileReport.EPPlusHelper.GetValue(worksheetName, riga, colonna2);

            if (value1 == null ||
                value2 == null ||
                !value1.ToString().Equals(value2.ToString(), StringComparison.InvariantCultureIgnoreCase))
            {
                // nome di reparto non valido, è presente in "Consumi spacchettati" ma manca in "Lista dati"
                throw new ManagedException(
                    tipologiaErrore: TipologiaErrori.DatoNonValido,
                    tipologiaCartella: TipologiaCartelle.ReportInput,
                    messaggioPerUtente: $"Il foglio '{worksheetName}' non ha le due nuove colonne 'Bdg Ore' e 'Bdg Lump' che devono essere presenti in posizione F e G",
                    worksheetName: worksheetName,
                    rigaCella: null,
                    colonnaCella: null,
                    dato: null,
                    percorsoFile: null
                    );
            }
        }



        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            //todo: rimuovere questo metodo, presente qui solo per assicurarsi che venga utilizzato un foglio Excel con la nuova struttura del foglio "consumi spacchettati"
            //ovvero con l'aggiunta delle due colonne "Bdg Ore" e "Bdg Lump" in posizione (F e G)
            tmpCheckVersioneConModificaInFoglioConsumiSpacchettatiCancellareAllaProssimaRealese(context);

            context.RigheTabellaConsumiSpacchettati = CalcolaConsumiSpacchettati(context.RepartiCensitiInReport, context.RigheSpese);
            context.UpdateReportsOutput.SettaTabellaConsumiSpacchettati(context.RigheTabellaConsumiSpacchettati);
            context.DebugInfoLogger.LogRigheTabellaConsumiSpacchettati(context.RigheTabellaConsumiSpacchettati);
            context.DebugInfoLogger.LogText("Calcolo righe tabella Consumi Spacchettati", context.RigheTabellaConsumiSpacchettati.Count);

            AggiornataTabellaConsumiSpacchettati(context.InfoFileReport, context.Configurazione, context.RepartiCensitiInReport, context.RigheTabellaConsumiSpacchettati);
            context.DebugInfoLogger.LogText("Aggiornamento tabella Consumi Spacchettati", "OK");

            return null;
        }

        private List<RigaTabellaConsumiSpacchettati> CalcolaConsumiSpacchettati(List<Reparto> repartiCensitiInReport, List<RigaSpese> righeSpese)
        {
            // vengono calcolati i totali di spesa per ogni reparto
            var consumiSpacchettati = new List<RigaTabellaConsumiSpacchettati>();
            foreach (var nomeReparto in repartiCensitiInReport.Select(_ => _.Nome).OrderBy(_ => _))
            {
                // Calcolo i totali per reparto
                // Spesa "Ore" - Actual
                var spesaAdOre_Actual_Euro = Math.Round(righeSpese.Where(_ => _.TipologiaDiSpesa == TipologieDiSpesa.AdOre && _.StatusSpesa == StatusSpesa.Actual && _.NomeReparto.Equals(nomeReparto, StringComparison.InvariantCultureIgnoreCase))
                                                         .Sum(_ => _.Spesa), Numbers.NumeroDecimaliImportiSpese);
                var spesaAdOre_Actual_Ore = Math.Round(righeSpese.Where(_ => _.TipologiaDiSpesa == TipologieDiSpesa.AdOre && _.StatusSpesa == StatusSpesa.Actual && _.NomeReparto.Equals(nomeReparto, StringComparison.InvariantCultureIgnoreCase))
                                                         .Sum(_ => _.Ore).Value, Numbers.NumeroDecimaliImportiSpese);
                // Spesa "Ore" - Commitment
                var spesaAdOre_Commitment_Euro = Math.Round(righeSpese.Where(_ => _.TipologiaDiSpesa == TipologieDiSpesa.AdOre && _.StatusSpesa == StatusSpesa.Commitment && _.NomeReparto.Equals(nomeReparto, StringComparison.InvariantCultureIgnoreCase))
                                                         .Sum(_ => _.Spesa), Numbers.NumeroDecimaliImportiSpese);
                var spesaAdOre_Commitment_Ore = Math.Round(righeSpese.Where(_ => _.TipologiaDiSpesa == TipologieDiSpesa.AdOre && _.StatusSpesa == StatusSpesa.Commitment && _.NomeReparto.Equals(nomeReparto, StringComparison.InvariantCultureIgnoreCase))
                                                         .Sum(_ => _.Ore).Value, Numbers.NumeroDecimaliImportiSpese);
                // Spesa "Lump sum" - Actual
                var spesaLumpSum_Actual = Math.Round(righeSpese.Where(_ => _.TipologiaDiSpesa == TipologieDiSpesa.LumpSum && _.StatusSpesa == StatusSpesa.Actual && _.NomeReparto.Equals(nomeReparto, StringComparison.InvariantCultureIgnoreCase))
                                                         .Sum(_ => _.Spesa), Numbers.NumeroDecimaliImportiSpese);
                // Spesa "Lump sum" - Commitment
                var spesaLumpSum_Commitment = Math.Round(righeSpese.Where(_ => _.TipologiaDiSpesa == TipologieDiSpesa.LumpSum && _.StatusSpesa == StatusSpesa.Commitment && _.NomeReparto.Equals(nomeReparto, StringComparison.InvariantCultureIgnoreCase))
                                                         .Sum(_ => _.Spesa), Numbers.NumeroDecimaliImportiSpese);

                var consumiPerReparto = new RigaTabellaConsumiSpacchettati(nomeReparto,
                                            spesaAdOre_Actual_Ore,
                                            spesaAdOre_Actual_Euro,
                                            spesaAdOre_Commitment_Euro,
                                            spesaAdOre_Commitment_Ore,
                                            spesaLumpSum_Actual,
                                            spesaLumpSum_Commitment);
                consumiSpacchettati.Add(consumiPerReparto);
            }
            return consumiSpacchettati;
        }
        private void AggiornataTabellaConsumiSpacchettati(InfoFileReport infoFileReport, Configurazione configurazione, List<Reparto> repartiCensitiInReport, List<RigaTabellaConsumiSpacchettati> consumiSpacchettati)
        {
            // IMPORTANTANTE: l'ordine dei nome reparti deve rimanere invariato, non posso quindi sostituire completamente il contenuto delle colonne
            // con la nuova tabella calcalata, ma devo cercare una per una le righe da aggiornare sulla base dei nomi dei reparti
            // N.B. Tutti i reparti presenti in "Consmumi spacchettati" deve essere necessariamente censito in "Lista dati", ma
            // solo i reparti censiti in "Lista dati" che abbiano anche delle spese devono necessariamente essere presenti in "Consmumi spacchettati" 
            var worksheetName = infoFileReport.WorksheetName_ConsumiSpacchettati; // "Consumi Spacchettati"
                                                                                  // il primo reparto si trova alla cella E16

            var rigaCorrente = configurazione.ConsumiSpacchettati_PrimaRigaConDati;

            // aggiorno le righe dei reparti presenti nella tabella in "Consmumi spacchettati" con i valori appena calcolati
            var repartiAggiornatiSuFoglioConsmumiSpacchettati = new List<string>();
            while (true)
            {
                var nomeReparto = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.ConsumiSpacchettati_ColonnaReparti);
                if (string.IsNullOrWhiteSpace(nomeReparto) || nomeReparto.Equals("Totale", StringComparison.CurrentCultureIgnoreCase))   // mi interrompo con l'ultima cella con valore nullo
                { break; }

                // verifico che il nome trovato sia presente tra quelli centisti in report (foglio "Lista dati")
                var repartoCensito = repartiCensitiInReport.SingleOrDefault(_ => _.Nome.Equals(nomeReparto, StringComparison.CurrentCultureIgnoreCase));
                if (repartoCensito == null)
                {
                    // nome di reparto non valido, è presente in "Consumi spacchettati" ma manca in "Lista dati"
                    throw new ManagedException(
                        tipologiaErrore: TipologiaErrori.DatoNonValido,
                        tipologiaCartella: TipologiaCartelle.ReportInput,
                        messaggioPerUtente: $"Il foglio '{worksheetName}' ha un reparto non correttamente censito nel foglio {infoFileReport.WorksheetName_ListaDati}')",
                        worksheetName: worksheetName,
                        rigaCella: rigaCorrente,
                        colonnaCella: configurazione.ConsumiSpacchettati_ColonnaReparti,
                        dato: nomeReparto,
                        percorsoFile: null
                        );
                }

                // uso Single() in quanto per costruzione la lista "consumiSpacchettati" deve aver necessariamente tutti i reparti con IsCensitoInReport = true
                var rigaConsumiSpacchettati = consumiSpacchettati.Single(_ => _.Reparto.Equals(nomeReparto, StringComparison.CurrentCultureIgnoreCase));
                infoFileReport.EPPlusHelper.SetValuesOnRow(worksheetName, rigaCorrente, configurazione.ConsumiSpacchettati_ColonnaOreActual,
                    rigaConsumiSpacchettati.SpesaAdOre_Actual_Ore,      // #1 "Ore Actual"
                    rigaConsumiSpacchettati.SpesaAdOre_Actual_Euro,     // #2 "Euro Actual"
                    rigaConsumiSpacchettati.SpesaAdOre_Commitment_Ore,  // #3 "Ore Committment"
                    rigaConsumiSpacchettati.SpesaAdOre_Commitment_Euro, // #4 "Euro Committment"
                    rigaConsumiSpacchettati.SpesaLumpSum_Actual,        // #5 "Lump Actual(FBL3N)"
                    rigaConsumiSpacchettati.SpesaLumpSum_Commitment     // #6 "Lump Committment(ME5A + ME2N)"
                    );

                repartiAggiornatiSuFoglioConsmumiSpacchettati.Add(repartoCensito.Nome);
                rigaCorrente++;
            }

            #region Verifico che il foglio "Consmumi Spacchettati" abbia tutti i reparti previsti
            // A questo punto ho verificato che tutti i reparti trovati in "Consmumi spacchettati" sono presenti tra quelli censiti.
            // Verifico quindi che tutti i reparti censiti (in "Lista dati") e che abbiano delle spese inputate a loro
            // siano effettivamente presenti in "Consmumi Spacchettati"
            foreach (var repartoCensitoInReport in repartiCensitiInReport)
            {
                if (!repartiAggiornatiSuFoglioConsmumiSpacchettati.Any(_ => _.Equals(repartoCensitoInReport.Nome, StringComparison.CurrentCultureIgnoreCase)))
                {
                    if (consumiSpacchettati.Single(_ => _.Reparto.Equals(repartoCensitoInReport.Nome, StringComparison.CurrentCultureIgnoreCase)).HasSpese)
                    {
                        // abbiamo un reparto censito in "Lista dati", non presenti in "Consmumi Spacchettati", ma con delle spese
                        // in questo caso il reparto dovrebbe essere nella lista di "Consmumi Spacchettati" ma risulta mancante.
                        throw new ManagedException(
                            tipologiaErrore: TipologiaErrori.DatoMancante,
                            tipologiaCartella: TipologiaCartelle.ReportInput,
                            messaggioPerUtente: $"Nel foglio '{worksheetName}' risulta mancare il reparto '{repartoCensitoInReport.Nome}')",
                            worksheetName: worksheetName,
                            rigaCella: null,
                            colonnaCella: null,
                            dato: repartoCensitoInReport.Nome,
                            nomeDatoErrore: NomiDatoErrore.NomeReparto,
                            percorsoFile: null
                            );
                    }
                }
            }
            #endregion
        }
    }
}