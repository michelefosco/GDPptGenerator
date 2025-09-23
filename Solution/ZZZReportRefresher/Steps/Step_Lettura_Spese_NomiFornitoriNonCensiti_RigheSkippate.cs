using ReportRefresher.Entities;
using ReportRefresher.Entities.Exceptions;
using ReportRefresher.Enums;
using System.Collections.Generic;
using System;
using System.Linq;
using ReportRefresher.Helpers;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Lettura spese e individuazione nomi dei Fornitori non censiti e righe skippate
    /// </summary>
    internal class Step_Lettura_Spese_NomiFornitoriNonCensiti_RigheSkippate : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            // Acquisizione Spese, nomi Fornitori non censiti, Righe skippate
            context.RigheSpese = GetListaSpese(context.DebugInfoLogger, context.InfoFileController, context.Configurazione, context.RepartiCensitiInReport, context.FornitoriCensitiInReport, context.UpdateReportsInput.AnnoCorrente, out context.FornitoriNonCensitiInReport, out context.RigheSpeseSkippate);

            // log delle righe di spesa
            context.UpdateReportsOutput.SettaRigheSpesa(context.RigheSpese);
            context.DebugInfoLogger.LogSpese(context.RigheSpese, "Tutte");

            // loggo le righe di spesa skippate
            context.UpdateReportsOutput.SettaRigheSpesaSkippate(context.RigheSpeseSkippate);
            context.DebugInfoLogger.LogSpesaSkippate(context.RigheSpeseSkippate);

            return null;
        }


        private List<RigaSpese> GetListaSpese(FileDebugHelper debugInfoLogger, InfoFileController infoFileController, Configurazione configurazione, List<Reparto> repartiCensitiInReport, List<FornitoreCensito> fornitoriCensiti, int annoCorrente, out List<FornitoreNonCensito> fornitoriNonCensitiInReport, out List<RigaSpeseSkippata> righeSpesaSkippate)
        {
            fornitoriNonCensitiInReport = new List<FornitoreNonCensito>();
            righeSpesaSkippate = new List<RigaSpeseSkippata>();

            var righeSpese = new List<RigaSpese>();

            // Spese "Ad ore" - Actual dal foglio ActualSoloCdc
            var costiSpesaAdOre_Actual = GetListaSpeseDaFoglio(
                    infoFileController: infoFileController,
                    worksheetName: infoFileController.WorksheetName_ActualSoloCdc,
                    tipologiaDiSpesa: TipologieDiSpesa.AdOre,
                    statusSpesa: StatusSpesa.Actual,
                    intestazionePrimaColonna: configurazione.ActualSoloCdc_IntestazionePrimaColonna,
                    colonnaCentroDiCosto: configurazione.ActualSoloCdc_ColonnaCentriDiCosto,
                    colonnaFornitore: configurazione.ActualSoloCdc_ColonnaFornitore,
                    colonnaFornitoreConDoppioValore: true,
                    colonnaSpesa: configurazione.ActualSoloCdc_ColonnaSpesa,
                    colonnaDataInizio: configurazione.ActualSoloCdc_ColonnaDataInizio,
                    colonnaDataFine: configurazione.ActualSoloCdc_ColonnaDataFine,
                    repartiCensiti: repartiCensitiInReport,
                    fornitoriCensiti: fornitoriCensiti,
                    fornitoriNonCensitiInReport: fornitoriNonCensitiInReport,
                    righeSpesaSkippate: righeSpesaSkippate,
                    annoCorrente: annoCorrente
                    );
            righeSpese.AddRange(costiSpesaAdOre_Actual);
            debugInfoLogger.LogSpese(costiSpesaAdOre_Actual, "AdOre_Actual");
            debugInfoLogger.LogText("Lettura spese ad ore Actual", costiSpesaAdOre_Actual.Count);

            // Spese "Ad ore" - Commitment dal foglio CommitmentSoloCdc
            var costiSpesaAdOre_Commitment = GetListaSpeseDaFoglio(
                    infoFileController: infoFileController,
                    worksheetName: infoFileController.WorksheetName_CommitmentSoloCdc,
                    tipologiaDiSpesa: TipologieDiSpesa.AdOre,
                    statusSpesa: StatusSpesa.Commitment,
                    intestazionePrimaColonna: configurazione.CommitmentSoloCdc_IntestazionePrimaColonna,
                    colonnaCentroDiCosto: configurazione.CommitmentSoloCdc_ColonnaCentriDiCosto,
                    colonnaFornitore: configurazione.CommitmentSoloCdc_ColonnaFornitore,
                    colonnaFornitoreConDoppioValore: false,
                    colonnaSpesa: configurazione.CommitmentSoloCdc_ColonnaSpesa,
                    colonnaDataInizio: configurazione.CommitmentSoloCdc_ColonnaDataInizio,
                    colonnaDataFine: configurazione.CommitmentSoloCdc_ColonnaDataFine,
                    repartiCensiti: repartiCensitiInReport,
                    fornitoriCensiti: fornitoriCensiti,
                    fornitoriNonCensitiInReport: fornitoriNonCensitiInReport,
                    righeSpesaSkippate: righeSpesaSkippate,
                    annoCorrente: annoCorrente
                    );
            righeSpese.AddRange(costiSpesaAdOre_Commitment);
            debugInfoLogger.LogSpese(costiSpesaAdOre_Commitment, "AdOre_Commitment");
            debugInfoLogger.LogText("Lettura spese ad ore commitment", costiSpesaAdOre_Commitment.Count);

            // Spese "Lump sum" - Actual dal foglio FBL3Nact
            var costiSpesaLumpSum_Actual = GetListaSpeseDaFoglio(
                    infoFileController: infoFileController,
                    worksheetName: infoFileController.WorksheetName_FBL3Nact,
                    tipologiaDiSpesa: TipologieDiSpesa.LumpSum,
                    statusSpesa: StatusSpesa.Actual,
                    intestazionePrimaColonna: configurazione.FBL3Nact_IntestazionePrimaColonna,
                    colonnaCentroDiCosto: configurazione.FBL3Nact_ColonnaCentriDiCosto,
                    colonnaFornitore: configurazione.FBL3Nact_ColonnaFornitore,
                    colonnaFornitoreConDoppioValore: true,
                    colonnaSpesa: configurazione.FBL3Nact_ColonnaSpesa,
                    colonnaDataInizio: configurazione.FBL3Nact_ColonnaDataInizio,
                    colonnaDataFine: configurazione.FBL3Nact_ColonnaDataFine,
                    repartiCensiti: repartiCensitiInReport,
                    fornitoriCensiti: fornitoriCensiti,
                    fornitoriNonCensitiInReport: fornitoriNonCensitiInReport,
                    righeSpesaSkippate: righeSpesaSkippate,
                    annoCorrente: annoCorrente
                    );
            righeSpese.AddRange(costiSpesaLumpSum_Actual);
            debugInfoLogger.LogSpese(costiSpesaLumpSum_Actual, "LumpSum_Actual");
            debugInfoLogger.LogText("Lettura spese lump sum Actual", costiSpesaLumpSum_Actual.Count);

            // Spese "Lump sum" - Commitment dal foglio ME5A
            var costiSpesaLumpSum_Commitment_ME5A = GetListaSpeseDaFoglio(
                    infoFileController: infoFileController,
                    worksheetName: infoFileController.WorksheetName_ME5A,
                    tipologiaDiSpesa: TipologieDiSpesa.LumpSum,
                    statusSpesa: StatusSpesa.Commitment,
                    intestazionePrimaColonna: configurazione.ME5A_IntestazionePrimaColonna,
                    colonnaCentroDiCosto: configurazione.ME5A_ColonnaCentriDiCosto,
                    colonnaFornitore: configurazione.ME5A_ColonnaFornitore,
                    colonnaFornitoreConDoppioValore: false,
                    colonnaSpesa: configurazione.ME5A_ColonnaSpesa,
                    colonnaDataInizio: configurazione.ME5A_ColonnaDataInizio,
                    colonnaDataFine: configurazione.ME5A_ColonnaDataFine,
                    repartiCensiti: repartiCensitiInReport,
                    fornitoriCensiti: fornitoriCensiti,
                    fornitoriNonCensitiInReport: fornitoriNonCensitiInReport,
                    righeSpesaSkippate: righeSpesaSkippate,
                    annoCorrente: annoCorrente
                    );
            righeSpese.AddRange(costiSpesaLumpSum_Commitment_ME5A);
            debugInfoLogger.LogSpese(costiSpesaLumpSum_Commitment_ME5A, "LumpSum_Commitment_ME5A");
            debugInfoLogger.LogText("Lettura spese lump sum commitment ME5A", costiSpesaLumpSum_Commitment_ME5A.Count);

            // Spese "Lump sum" - Commitment dal foglio ME2N
            var costiSpesaLumpSum_Commitment_ME2N = GetListaSpeseDaFoglio(
                    infoFileController: infoFileController,
                    worksheetName: infoFileController.WorksheetName_ME2N,
                    tipologiaDiSpesa: TipologieDiSpesa.LumpSum,
                    statusSpesa: StatusSpesa.Commitment,
                    intestazionePrimaColonna: configurazione.ME2N_IntestazionePrimaColonna,
                    colonnaCentroDiCosto: configurazione.ME2N_ColonnaCentriDiCosto,
                    colonnaFornitore: configurazione.ME2N_ColonnaFornitore,
                    colonnaFornitoreConDoppioValore: true,
                    colonnaSpesa: configurazione.ME2N_ColonnaSpesa,
                    colonnaDataInizio: configurazione.ME2N_ColonnaDataInizio,
                    colonnaDataFine: configurazione.ME2N_ColonnaDataFine,
                    repartiCensiti: repartiCensitiInReport,
                    fornitoriCensiti: fornitoriCensiti,
                    fornitoriNonCensitiInReport: fornitoriNonCensitiInReport,
                    righeSpesaSkippate: righeSpesaSkippate,
                    annoCorrente: annoCorrente
                    );
            righeSpese.AddRange(costiSpesaLumpSum_Commitment_ME2N);
            debugInfoLogger.LogSpese(costiSpesaLumpSum_Commitment_ME2N, "LumpSum_Commitment_ME2N");
            debugInfoLogger.LogText("Lettura spese lump sum commitment ME2N", costiSpesaLumpSum_Commitment_ME2N.Count);

            return righeSpese;
        }
        private List<RigaSpese> GetListaSpeseDaFoglio(InfoFileController infoFileController, string worksheetName, TipologieDiSpesa tipologiaDiSpesa, StatusSpesa statusSpesa, string intestazionePrimaColonna, int colonnaCentroDiCosto, int colonnaFornitore, bool colonnaFornitoreConDoppioValore, int colonnaSpesa, int colonnaDataInizio, int colonnaDataFine, List<Reparto> repartiCensiti, List<FornitoreCensito> fornitoriCensiti, List<FornitoreNonCensito> fornitoriNonCensitiInReport, List<RigaSpeseSkippata> righeSpesaSkippate, int annoCorrente)
        {
            var listaRigaSpese = new List<RigaSpese>();

            // determino prima e ultima riga da scansionare per le spese
            var rigaCorrente = infoFileController.EPPlusHelper.GetFirstRowWithSpecificValue(worksheetName, 1, 1, intestazionePrimaColonna);
            if (rigaCorrente < 1)
            {
                throw new ManagedException(
                            tipologiaErrore: TipologiaErrori.DatoMancante,
                            tipologiaCartella: TipologiaCartelle.Controller,
                            worksheetName: worksheetName,
                            rigaCella: null,
                            colonnaCella: null,
                            nomeDatoErrore: NomiDatoErrore.IntestazioneDati,
                            dato: null,
                            messaggioPerUtente: $"Impossibile individurare la riga intestazione (riga con prima cella contenente il testo '{intestazionePrimaColonna}')"
                        );
            }

            var ultimaRigaUsataNelFoglio = infoFileController.EPPlusHelper.GetRowsLimit(worksheetName);

            while (rigaCorrente < ultimaRigaUsataNelFoglio)
            {
                rigaCorrente++; // avanza di una riga

                // Lettura dei dati dalle colonne del foglio
                var centroDiCosto = infoFileController.EPPlusHelper.GetString(worksheetName, rigaCorrente, colonnaCentroDiCosto);
                var stringaFornitore = infoFileController.EPPlusHelper.GetString(worksheetName, rigaCorrente, colonnaFornitore);
                var objectSpesa = infoFileController.EPPlusHelper.GetValue(worksheetName, rigaCorrente, colonnaSpesa);
                var objectDataInizio = infoFileController.EPPlusHelper.GetValue(worksheetName, rigaCorrente, colonnaDataInizio);
                var objectDataFine = infoFileController.EPPlusHelper.GetValue(worksheetName, rigaCorrente, colonnaDataFine);

                // Se tutte i valori sono nulli ho raggiunto la fine della tabella, mi fermo
                if (centroDiCosto == null && objectSpesa == null && stringaFornitore == null && objectDataInizio == null && objectDataFine == null)
                { break; }

                // Il centro di costo deve appartenere ad uno dei reparti censiti
                ValidaCentoDiCosto(worksheetName, rigaCorrente, colonnaCentroDiCosto, centroDiCosto, repartiCensiti, out string nomeReparto);

                // Le righe con un fornitore che inizia con il carattere '#' vanno skippata ma NON è considerato un errore
                if (stringaFornitore != null && stringaFornitore.TrimStart().StartsWith("#"))
                {                    
                    righeSpesaSkippate.Add(new RigaSpeseSkippata(worksheetName, rigaCorrente, colonnaFornitore, nomeReparto, stringaFornitore));
                    continue;
                }
                // Il nome del fornitore deve essere tra quelli censiti
                ValidaFornitore(worksheetName, rigaCorrente, colonnaFornitore, stringaFornitore, colonnaFornitoreConDoppioValore, fornitoriCensiti, fornitoriNonCensitiInReport, nomeReparto, tipologiaDiSpesa, out FornitoreCensito fornitore);

                // La spesa deve avere un valore corretto (potrebbe contenere nulli o valori come "#n/a")
                if (!IsSpesaValida(objectSpesa, out double? spesa))
                {
                    // ho un valore di spesa non valido, salto la riga e la registro tra quelle skippate
                    righeSpesaSkippate.Add(new RigaSpeseSkippata(worksheetName, rigaCorrente, colonnaSpesa, nomeReparto, objectSpesa));
                    continue;
                }

                // Valori valorizzati solo in caso di spese di tipologia "Ad ore"
                double? oreSvolte = null;
                double? costoOrarioApplicato = null;
                if (tipologiaDiSpesa == TipologieDiSpesa.AdOre)
                {
                    // per le spese di Tipologia "Ad ore" calcolo anche le ore svolte
                    // il fornitore deve avere un costo orario (eventualmente specifico per il reparto) configurato in modo corretto (>0)
                    ValidaCostoOrarioPerFornitore(worksheetName, rigaCorrente, colonnaSpesa, fornitore, nomeReparto, out double costoOrarioFornitorePerReparto);
                    oreSvolte = spesa.Value / costoOrarioFornitorePerReparto;
                    costoOrarioApplicato = costoOrarioFornitorePerReparto;
                }

                // verifico le eventuali date di Inizio e Fine del periodo di spesa
                ValidaDataInizioPeriodoSpesa(worksheetName, rigaCorrente, colonnaDataInizio, objectDataInizio, out DateTime dataInizioValidata);
                ValidaDataFinePeriodoSpesa(worksheetName, rigaCorrente, colonnaDataInizio, objectDataFine, dataInizioValidata, out DateTime dataFineValidata);

                // aggiunta riga alla lista
                var rigaCosti = new RigaSpese(centroDiCosto, nomeReparto, fornitore, tipologiaDiSpesa, statusSpesa, spesa.Value, oreSvolte, costoOrarioApplicato, dataInizioValidata, dataFineValidata);
                listaRigaSpese.Add(rigaCosti);

                // Valuto se il fornitore era marcato come "PresenteSoloInListaDati".
                // In questo caso, avendo delle spese ad esso associtate, dovrà essere inserito ovunque nel report
                if (fornitore.PresenteSoloInListaDati)
                {
                    fornitore.Setta_DeveEsserePresenteNeiReport();
                }
            }
            return listaRigaSpese;
        }

        #region Validazione dei dati letti
        private void ValidaCentoDiCosto(string worksheetName, int rigaCella, int colonnaCella, string centroDiCosto, List<Reparto> repartiCensiti, out string nomeReparto)
        {
            if (centroDiCosto == null)
            {
                throw new ManagedException(
                            tipologiaErrore: TipologiaErrori.DatoMancante,
                            tipologiaCartella: TipologiaCartelle.Controller,
                            worksheetName: worksheetName,
                            rigaCella: rigaCella,
                            colonnaCella: colonnaCella,
                            nomeDatoErrore: NomiDatoErrore.CentroDiCosto,
                            dato: null
                            );
            }

            var reparto = repartiCensiti.SingleOrDefault(r => r.CentriDiCosto.Any(cdc => cdc.Equals(centroDiCosto, StringComparison.InvariantCultureIgnoreCase)));
            if (reparto == null)
            {
                throw new ManagedException(
                            tipologiaErrore: TipologiaErrori.DatoNonValido,
                            tipologiaCartella: TipologiaCartelle.Controller,
                            worksheetName: worksheetName,
                            rigaCella: rigaCella,
                            colonnaCella: colonnaCella,
                            nomeDatoErrore: NomiDatoErrore.CentroDiCosto,
                            dato: centroDiCosto,
                            messaggioPerUtente: $"Nel foglio '{worksheetName}' nella cella '{(ColumnIDS)colonnaCella}{rigaCella}' è stato trovato un centro di costo non censito ('{centroDiCosto}').\n(Probabilmente è necessario aggiungere manualmente il nuovo reparto nei report. Contattare Perrino)"
                            );
            }

            nomeReparto = reparto.Nome;
        }
        private bool IsSpesaValida(object objectSpesa, out double? spesa)
        {
            spesa = objectSpesa as double?;
            return spesa.HasValue;
        }
        private void ValidaFornitore(string worksheetName, int rigaCella, int colonnaCella, object objectFornitore, bool colonnaConDoppioValore, List<FornitoreCensito> fornitoriCensiti, List<FornitoreNonCensito> fornitoriNonCensitiInReport, string nomeReparto, TipologieDiSpesa tipologiaDiSpesa, out FornitoreCensito fornitore)
        {
            if (objectFornitore == null || string.IsNullOrWhiteSpace(objectFornitore.ToString()))
            {
                throw new ManagedException(
                            tipologiaErrore: TipologiaErrori.DatoMancante,
                            tipologiaCartella: TipologiaCartelle.Controller,
                            worksheetName: worksheetName,
                            rigaCella: rigaCella,
                            colonnaCella: colonnaCella,
                            nomeDatoErrore: NomiDatoErrore.NomeFornitore,
                            dato: null
                            );
            }

            string nomeFornitore;
            if (colonnaConDoppioValore)
            {
                if (!ValueHelper.CastNomeFornitoreDaColonnaConDoppioValore(objectFornitore, out nomeFornitore))
                {
                    throw new ManagedException(
                        tipologiaErrore: TipologiaErrori.DatoNonValido,
                        tipologiaCartella: TipologiaCartelle.Controller,
                        worksheetName: worksheetName,
                        rigaCella: rigaCella,
                        colonnaCella: colonnaCella,
                        dato: objectFornitore.ToString(),
                        nomeDatoErrore: NomiDatoErrore.NomeFornitore
                        );
                }
            }
            else
            {
                nomeFornitore = objectFornitore.ToString().Trim();
            }

            var nomeFornitoreDaCercare = nomeFornitore;
            fornitore = fornitoriCensiti.SingleOrDefault(fc => fc.HasThisName(nomeFornitoreDaCercare));
            if (fornitore == null)
            {
                // se non trovo il fornitore tra quelli censiti non sollevo un eccezione ma mi limito a segnare il nome del fornitore
                // trovato nella lista dei fornitore non censiti. In questo modo posso adare avanti con  la scansione di tutti i fogli
                // di spesa dove potrei trovare più di un nuovo fornitore
                var fornitoriNonCensitoInReport = fornitoriNonCensitiInReport.SingleOrDefault(fc => fc.NomeSuController.Equals(nomeFornitoreDaCercare, StringComparison.InvariantCultureIgnoreCase));

                // mi assicuro di aggiungere una sola volta i nomi dei nuovi fornitori trovati
                if (fornitoriNonCensitoInReport == null)
                {
                    var nuovoFornitoriNonCensitoInReport = new FornitoreNonCensito(nomeFornitoreDaCercare);
                    nuovoFornitoriNonCensitoInReport.AssociaFoglioAndRepartoAlNomeDelNuovoFornitore(worksheetName, nomeReparto, tipologiaDiSpesa);
                    fornitoriNonCensitiInReport.Add(nuovoFornitoriNonCensitoInReport);
                }
                else
                {
                    // il nuovo fornitore era già stato incontrato tra le spese
                    // in questo caso aggiorno solo la lista dei nomi dei fogli in cui è stato trovato
                    fornitoriNonCensitoInReport.AssociaFoglioAndRepartoAlNomeDelNuovoFornitore(worksheetName, nomeReparto, tipologiaDiSpesa);
                }
                fornitore = new FornitoreCensito(siglaInReport: nomeFornitoreDaCercare,
                                                        nomeSuController: nomeFornitoreDaCercare,
                                                        presenteSoloInListaDati: false);
            }
        }
        private void ValidaCostoOrarioPerFornitore(string worksheetName, int rigaCella, int colonnaCella, FornitoreCensito fornitore, string nomeReparto, out double costoOrarioFornitorePerReparto)
        {
            if (!fornitore.HasCostoOrarioSettato)
            {
                // Caso fornitore non ancora censito
                // Siamo di fronte ad una riga di spesa per un nuovo fornitore non ancora censito, di cui quindi non conosciamo i costi orari.
                // Restituiamo 0 consapevoli del fatto che i dati calcolati per le ore non verranno utlizzanti, in quanto
                // il processo terminerà con esito: EsitiFinali.FornitoriDaAggiungere
                costoOrarioFornitorePerReparto = 0;
                return;
            }

            costoOrarioFornitorePerReparto = fornitore.GetCostoOrarioPerReparto(nomeReparto);
            if (costoOrarioFornitorePerReparto <= 0)
            {
                throw new ManagedException(
                         tipologiaErrore: TipologiaErrori.DatoNonValido,
                         tipologiaCartella: TipologiaCartelle.ReportInput,
                         // Messaggio concordato con Francesco, non cambiare.
                         messaggioPerUtente: $"Il fornitore '{fornitore.SiglaInReport}' che ha una o più voci di spesa ad ORE (foglio '{worksheetName}'), ha un costo orario pari a zero. Questo rende impossibile il calcolo delle ore(Ore = spesa / costo orario) in quanto si avrebbe una divisione per zero.\nInserire il costo orario nel foglio ‘Lista Dati’ del file Excel e ripetere la procedura.",
                         worksheetName: worksheetName,
                         rigaCella: rigaCella,
                         colonnaCella: colonnaCella,
                         nomeDatoErrore: NomiDatoErrore.CostoOrarioFornitore,
                         dato: null
                         );
            }
        }
        private void ValidaDataInizioPeriodoSpesa(string worksheetName, int rigaCella, int colonnaCella, object objectData, out DateTime dataInizioValidata)
        {
            if (objectData == null)
            {
                throw new ManagedException(
                            tipologiaErrore: TipologiaErrori.DatoMancante,
                            tipologiaCartella: TipologiaCartelle.Controller,
                            worksheetName: worksheetName,
                            rigaCella: rigaCella,
                            colonnaCella: colonnaCella,
                            nomeDatoErrore: NomiDatoErrore.DataInizioPeriodoSpesa,
                            dato: null
                            );
            }

            // tento il cast dell'object in data
            var dataValidataTest = objectData as DateTime?;
            if (!dataValidataTest.HasValue)
            {
                throw new ManagedException(
                            tipologiaErrore: TipologiaErrori.DatoNonValido,
                            tipologiaCartella: TipologiaCartelle.Controller,
                            worksheetName: worksheetName,
                            rigaCella: rigaCella,
                            colonnaCella: colonnaCella,
                            nomeDatoErrore: NomiDatoErrore.DataInizioPeriodoSpesa,
                            dato: objectData.ToString()
                            );
            }

            // cast riuscito
            dataInizioValidata = dataValidataTest.Value;
        }
        private void ValidaDataFinePeriodoSpesa(string worksheetName, int rigaCella, int colonnaCella, object objectData, DateTime dataInizio, out DateTime dataFineValidata)
        {
            if (objectData == null)
            {
                throw new ManagedException(
                            tipologiaErrore: TipologiaErrori.DatoMancante,
                            tipologiaCartella: TipologiaCartelle.Controller,
                            worksheetName: worksheetName,
                            rigaCella: rigaCella,
                            colonnaCella: colonnaCella,
                            nomeDatoErrore: NomiDatoErrore.DataFinePeriodoSpesa,
                            dato: null
                            );
            }

            // tento il cast dell'object in data
            var dataValidataTest = objectData as DateTime?;
            if (!dataValidataTest.HasValue)
            {
                throw new ManagedException(
                            tipologiaErrore: TipologiaErrori.DatoNonValido,
                            tipologiaCartella: TipologiaCartelle.Controller,
                            worksheetName: worksheetName,
                            rigaCella: rigaCella,
                            colonnaCella: colonnaCella,
                            nomeDatoErrore: NomiDatoErrore.DataFinePeriodoSpesa,
                            dato: objectData.ToString()
                            );
            }

            // cast riuscito
            dataFineValidata = dataValidataTest.Value;

            // verifico che la data di fine si maggiore o uguale alla data di inizio
            if (dataFineValidata < dataInizio)
            {
                throw new ManagedException(
                        tipologiaErrore: TipologiaErrori.DatoNonValido,
                        tipologiaCartella: TipologiaCartelle.Controller,
                         worksheetName: worksheetName,
                        rigaCella: rigaCella,
                        colonnaCella: colonnaCella,
                        nomeDatoErrore: NomiDatoErrore.DataFinePeriodoSpesa,
                        dato: dataFineValidata.ToShortDateString(),
                        messaggioPerUtente: "La data di fine del periodo di riferimento della spesa deve essere successiva alla data di inizio"
                        );
            }
        }
        #endregion
    }
}