using ReportRefresher.Entities;
using ReportRefresher.Entities.Exceptions;
using ReportRefresher.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Lettura info reparti censiti nel file controller
    /// </summary>
    internal class Step_Lettura_RepartiCensitiSuController : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            // Creazione lista reparti dal file "Controller"
            context.RepartiCensitiInController = GetRepartiCensitiSuController(context.InfoFileController, context.Configurazione);
            context.DebugInfoLogger.LogText("Lettura reparti censiti nel file 'Controller'", context.RepartiCensitiInController.Count);

            // Merge di alcuni reparti sotto un unico reparto
            MergeCentroDiCostiPerReparti(context.RepartiCensitiInController, context.Configurazione);
            context.DebugInfoLogger.LogText("Merge dei codici di alcuni reparti censiti nel file 'Controller'", "OK");
            context.UpdateReportsOutput.SettaRepartiCensitiController(context.RepartiCensitiInController);
            context.DebugInfoLogger.LogRepartiCensitiSuController(context.RepartiCensitiInController);

            // Creazione lista reparti dal file "Report" foglio "Lista dati" validando i dati con quelli trovati nel file Cotroller foglio "Recap" 
            context.RepartiCensitiInReport = GetRepartiCensitiInReport(context.InfoFileReport, context.Configurazione, context.RepartiCensitiInController);
            context.DebugInfoLogger.LogText("Lettura reparti censiti nel file 'Report'", context.RepartiCensitiInReport.Count);
            context.UpdateReportsOutput.SettaRepartiCensitiInReport(context.RepartiCensitiInReport);
            context.DebugInfoLogger.LogrepartiCensitiSuReport(context.RepartiCensitiInReport);

            return null;
        }

        /// <summary>
        /// Legge l'elenco dei reparti e dei codici dei centri di costo ad essi associati
        /// </summary>
        /// <param name="InfoFileController"></param>
        private List<Reparto> GetRepartiCensitiSuController(InfoFileController InfoFileController, Configurazione configurazione)
        {
            var worksheetName = InfoFileController.WorksheetName_Recap;

            var listaReparti = new List<Reparto>();

            var ultimaRigaUsataNelFoglio = InfoFileController.EPPlusHelper.GetRowsLimit(worksheetName);
            var rigaCorrente = configurazione.Recap_Riga_PrimaConDati;

            var primoElementoDiSegmento = true;
            var cdcLettiPerSegmentoCorrente = new List<string>();
            var nomeRepartoDelSegmentoCorrente = "";
            while (rigaCorrente <= ultimaRigaUsataNelFoglio)
            {
                var valueCdc = InfoFileController.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.Recap_ColonnaCentriDiCosto);

                if (!string.IsNullOrWhiteSpace(valueCdc))
                {
                    // i centri di costo devono essere univoci in modo da essere assegnati univocamente ad un unico reparto
                    if (listaReparti.Any(r => r.CentriDiCosto.Any(cdc => cdc.Equals(valueCdc, StringComparison.InvariantCultureIgnoreCase))))
                    {
                        //TEST: InputFile_Controller.CDC_NonUnivoco
                        //InputExportController_KO_CentroDiCostoNonUnivoco.xlsx
                        throw new ManagedException(
                            tipologiaErrore: TipologiaErrori.DatoNonUnivoco,
                            tipologiaCartella: TipologiaCartelle.Controller,
                            worksheetName: worksheetName,
                            rigaCella: rigaCorrente,
                            colonnaCella: configurazione.Recap_ColonnaCentriDiCosto,
                            nomeDatoErrore: NomiDatoErrore.CentroDiCosto,
                            dato: valueCdc,
                            percorsoFile: null
                            );
                    }

                    // registro il codice appena letto
                    cdcLettiPerSegmentoCorrente.Add(valueCdc);

                    if (primoElementoDiSegmento)
                    {
                        // se sono sul primo elemento del segmento devo necessariamente avere anche il nome del reparto
                        nomeRepartoDelSegmentoCorrente = InfoFileController.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.Recap_ColonnaReparti);

                        // il reparto deve essere presente all'inizio di un segmento di centri di costo
                        if (string.IsNullOrWhiteSpace(nomeRepartoDelSegmentoCorrente))
                        {
                            throw new ManagedException(
                            tipologiaErrore: TipologiaErrori.DatoMancante,
                            tipologiaCartella: TipologiaCartelle.Controller,
                            worksheetName: worksheetName,
                            colonnaCella: configurazione.Recap_ColonnaReparti,
                            rigaCella: rigaCorrente,
                            nomeDatoErrore: NomiDatoErrore.NomeReparto,
                            dato: null
                            );
                        }
                    }

                    // ho letto un valore di cdc --> non sono più al primoElementoDiSegmento
                    primoElementoDiSegmento = false;
                }
                else
                {
                    // valueCdc = "" significa che sono in presenza di un'interruzione di segmento
                    // determino se è appena terminato un segmento valido con un reparto e almeno un Centro di costo
                    if (!string.IsNullOrEmpty(nomeRepartoDelSegmentoCorrente) && cdcLettiPerSegmentoCorrente.Any())
                    {
                        // decido se aggiungere i codici ad un reparto già censito o aggiungee un nuovo reparto alla lista
                        var reparto = listaReparti.SingleOrDefault(_ => _.Nome.Equals(nomeRepartoDelSegmentoCorrente, StringComparison.InvariantCultureIgnoreCase));
                        if (reparto == null)
                        {
                            reparto = new Reparto(nomeRepartoDelSegmentoCorrente);
                            listaReparti.Add(reparto);
                        }
                        else
                        {
                            // Questo scenario non rappresenta un errore in quanto nel file controller arrivato da GD il nome di alcuni
                            // reparti è indicato più di una volta nel foglio "RECAP" (vedi Gamberini o Monti) ma in Aree diverse
                        }
                        reparto.AggiungiCentriDiCosto(cdcLettiPerSegmentoCorrente);
                    }

                    // dato che sono in presenza di un'interruzione di segmento (caso cdc: null), azzero le info relative al segmento
                    // appena terminato per predisporto al prossimo
                    primoElementoDiSegmento = true;
                    nomeRepartoDelSegmentoCorrente = "";
                    cdcLettiPerSegmentoCorrente = new List<string>();
                }

                // vado alla riga successiva
                rigaCorrente++;
            }

            return listaReparti;
        }
        private  void MergeCentroDiCostiPerReparti(List<Reparto> repartiCensiti, Configurazione configurazione)
        {
            //FRANCESCO DICE: i seguenti reparti: Milandri, Testoni, Gianese, Moretti devono fare tutti capo al reparto Lanzarini. Cioè per intenderci: dove trovi uno di quei nomi, puoi censirli tutti come se fossero Lanzarini. È un eccezione unica.
            foreach (var mergeReparti in configurazione.Recap_ListaMergeReparti)
            {
                var valoriSplittati = mergeReparti.Split('|');
                if (valoriSplittati.Length == 2)
                {
                    string nomeRepartoFrom = valoriSplittati[0].Trim();
                    string nomeRepartoTo = valoriSplittati[1].Trim();
                    MergeCentriDiCostoReparto(repartiCensiti, nomeRepartoFrom, nomeRepartoTo);
                }
            }
        }
        private  void MergeCentriDiCostoReparto(List<Reparto> listaReparti, string nomeRepartoFrom, string nomeRepartoTo)
        {
            var repartoFrom = listaReparti.SingleOrDefault(_ => _.Nome.Equals(nomeRepartoFrom, StringComparison.InvariantCultureIgnoreCase));
            var repartoTo = listaReparti.SingleOrDefault(_ => _.Nome.Equals(nomeRepartoTo, StringComparison.InvariantCultureIgnoreCase));
            if (repartoFrom != null && repartoTo != null)
            {
                repartoTo.AggiungiCentriDiCosto(repartoFrom.CentriDiCosto);
                repartoFrom.AzzeraCentriDiCosto();
            }
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

                // individuo il reparto già letto dal foglio "Recap" il quale ha anche le iniziali del nome.
                // prendo quindi il reparto il cui nome inizia con il nome del reparto presente in "Lista dati" e per evitiare situazioni
                // in cui ci sia più di un reparto il cui nome comincia allo stesso modo, prendo quello con il nome più lungo (match più accurato)
                var repartoCensitiInController = repartiCensitiInController.Where(_ => _.Nome.ToUpper().StartsWith(nomeRepartoSuListaDati.ToUpper())).OrderBy(_ => _.Nome.Length).LastOrDefault();
                if (repartoCensitiInController != null)
                {
                    // creo un nuovo reparto con il nome trovato in Lista dati (senza iniziali del nome) e i centri di costo come letti dal file controller
                    var repartoCensitoInReport = new Reparto(nomeRepartoSuListaDati);
                    repartoCensitoInReport.AggiungiCentriDiCosto(repartoCensitiInController.CentriDiCosto);
                    repartiCensitoInReport.Add(repartoCensitoInReport);
                }
                else
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
                colonnaCorrente++;
            }
            return repartiCensitoInReport;
        }
     }
}