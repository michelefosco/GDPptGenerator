using ReportRefresher.Entities;
using ReportRefresher.Entities.Exceptions;
using ReportRefresher.Enums;
using System.Collections.Generic;
using System.Linq;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Lettura info reparti censiti
    /// </summary>
    internal class Step_Lettura_RepartiCensitiInReport : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            // Creazione lista reparti dal file "Report" foglio "Lista dati" validando i dati con quelli trovati nel file Cotroller foglio "Recap" 
            context.RepartiCensitiInReport = GetRepartiCensitiInReport(context.InfoFileReport, context.Configurazione, context.RepartiCensitiInController);
            context.DebugInfoLogger.LogText("Lettura reparti censiti nel file 'Report'", context.RepartiCensitiInReport.Count);
            context.UpdateReportsOutput.SettaRepartiCensitiInReport(context.RepartiCensitiInReport);
            context.DebugInfoLogger.LogrepartiCensitiSuReport(context.RepartiCensitiInReport);

            return null;
        }
        private List<Reparto> GetRepartiCensitiInReport(InfoFileReport infoFileReport, Configurazione configurazione, List<Reparto> repartiCensitiInController)
        {
            var worksheetName = infoFileReport.WorksheetName_ListaDati; // Lista dati
            var repartiCensitoInReport = new List<Reparto>();

            // scorro verso destra le colonne della tabella contenente i nomi dei reparti
            var rigaReparti = configurazione.ListaDati_RigaReparti;
            var colonnaCorrente = configurazione.ListaDati_PrimaColonnaReparti;
            while (true)
            {
                var nomeRepartoSuListaDati = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaReparti, colonnaCorrente);
                if (string.IsNullOrWhiteSpace(nomeRepartoSuListaDati)) // mi interrompo con l'ultima cella con eventuale valore blank
                { break; }

                // creo un nuovo reparto con il nome trovato in Lista dati (senza iniziali del nome) e i centri di costo come letti dal file controller
                var repartoCensitoInReport = new Reparto(nomeRepartoSuListaDati);

                // Se repartiCensitiInController è valorizzato lo uso per verificare
                // che i reparti trovati in Report siano presenti anche su controller
                if (repartiCensitiInController != null)
                {
                    // individuo il reparto già letto dal foglio "Recap" il quale ha anche le iniziali del nome.
                    // prendo quindi il reparto il cui nome inizia con il nome del reparto presente in "Lista dati" e per evitiare situazioni
                    // in cui ci sia più di un reparto il cui nome comincia allo stesso modo, prendo quello con il nome più lungo (match più accurato)
                    var repartoCensitiInController = repartiCensitiInController.Where(_ => _.Nome.ToUpper().StartsWith(nomeRepartoSuListaDati.ToUpper())).OrderBy(_ => _.Nome.Length).LastOrDefault();
                    if (repartoCensitiInController == null)
                    {
                        // FRANCESCO dice:
                        // In Lista Dati ci sono meno reparti poiché ai fini dei report non occorre il censimento di tutti i reparti di Ufficio Tecnico.
                        // Nelle voci di spesa (ovvero fogli ‘act_solo_cdc’, ‘FBL3N’, ‘committment_solo_cdc’, ‘ME5A’, ‘ME2N’) ci sono solo alcuni dei reparti.
                        // Il codice potrebbe censirli tutti, e poi usarli come Dizionario, prendendo solo quelli che appaiono nelle voci di spesa.
                        // Bisogna segnalare all’utente solo se tra le voci di spesa del file dei controller(non nel foglio RECAP) appare un centro di costo
                        // afferente ad un reparto NON presente in lista dati. In quel caso toccherà a noi aggiungerlo nelle tabelle dei reports.
                        throw new ManagedException(
                            tipologiaErrore: TipologiaErrori.DatoNonValido,
                            tipologiaCartella: TipologiaCartelle.ReportInput,
                            messaggioPerUtente: $"Nel file Report nel foglio '{worksheetName}' è presente il reparto '{nomeRepartoSuListaDati}' il quale non tra quelli censiti sul file controller).",
                            worksheetName: worksheetName,
                            rigaCella: rigaReparti,
                            colonnaCella: colonnaCorrente,
                            dato: nomeRepartoSuListaDati,
                            percorsoFile: null
                            );
                    }

                    // aggiungo al reparto le informazioni sui centri di costo come recuperate dalle informazioni dei reparti lette dal file controller
                    repartoCensitoInReport.AggiungiCentriDiCosto(repartoCensitiInController.CentriDiCosto);
                }             
                
                // Aggiungo il reparto appena trovato alla lista
                repartiCensitoInReport.Add(repartoCensitoInReport);

                colonnaCorrente++;
            }
            return repartiCensitoInReport;
        }
    }
}