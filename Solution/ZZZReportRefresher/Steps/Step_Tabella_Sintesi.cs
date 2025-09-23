using ReportRefresher.Entities;
using ReportRefresher.Entities.Exceptions;
using ReportRefresher.Enums;
using System.Collections.Generic;
using System;
using ReportRefresher.Helpers;
using System.Linq;
using System.Drawing;
using ReportRefresher.Constants;

namespace ReportRefresher.Steps
{
    //todo: rendere internal ma accessibile ai test
    /// <summary>
    /// Gestione tabella "Sintesi"
    /// </summary>
    public class Step_Tabella_Sintesi : Step_Base
    {
        Color BACKGROUND_COLOR_PAST_MONTHS = Color.FromArgb(0, 255, 255, 204);        //FF 70 30 A0
        Color BACKGROUND_COLOR_FUTURE_MONTHS = Color.White;
        Color BACKGROUND_COLOR_EXTRABUDGET = Color.Red;

        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            #region Operazioni preliminari
            // Determino la lista delle tuple necessarie per coprire tutte le combinazioni di "Fornitore/Reparto/Tipo spesa" presenti nella tabella avvanzamento
            var tupleSpesePerAnno = CreaSetMinimoTupleSpese(context.RigheTabellaAvanzamento);
            //var tupleSpesePerAnno = CreaSetMinimoTupleSpese_V2(context.RigheSpese);
            context.UpdateReportsOutput.SettaTabellaSintesi_SetMinimoRigheNecessarie(tupleSpesePerAnno);
            context.DebugInfoLogger.LogRigheTabellaSintesi_Caclcolate(tupleSpesePerAnno);
            context.DebugInfoLogger.LogText("Calcolo del set di minimo di tuple per la tabella Sintesi", tupleSpesePerAnno.Count);

            // Vengono calcolati i totali delle spese mensili per le tuple
            AssegnaTotaliSpeseMensiliAlleTuple(tupleSpesePerAnno, context.RigheSpese, context.UpdateReportsInput.AnnoCorrente);
            context.DebugInfoLogger.LogRigheTabellaSintesi_DatiSpeseConfermate(tupleSpesePerAnno);
            context.DebugInfoLogger.LogText("Ricavati i valori di spesa per le tuple", "OK");
            #endregion

            #region Operazioni su tabella "Sintesi"
            // Lettura dei valori attuali in tabella
            var righeTabellaSintesi = GetRigheTabellaSintesi(context.InfoFileReport, context.Configurazione, context.RepartiCensitiInReport, context.FornitoriCensitiInReport);
            context.DebugInfoLogger.LogRigheTabellaSintesi_CurrentValues(righeTabellaSintesi);
            context.DebugInfoLogger.LogText("Lettura righe attualmente presenti tabella Sintesi", righeTabellaSintesi.Count);

            // Creo una lista di righe che rappresenta le tupleSpese per cui manca almeno una riga in tabella Sintesi
            var righeMancantiTabellaSintesi = IndividuaRigheMancantiTabellaSintesi(righeTabellaSintesi, tupleSpesePerAnno, context.FornitoriCensitiInReport);
            context.UpdateReportsOutput.SettaTabellaSintesi_RigheMancanti(righeMancantiTabellaSintesi);
            context.DebugInfoLogger.LogRigheTabellaSintesi_Mancanti(righeMancantiTabellaSintesi);
            context.DebugInfoLogger.LogText("Individuazione delle righe mancanti nella tabella Sintesi", righeMancantiTabellaSintesi.Count);

            // Aggiungo le righe mancanti
            AggiungiRigheMancantiTabelleSintesi(context.InfoFileReport, context.Configurazione, righeTabellaSintesi, righeMancantiTabellaSintesi);
            context.DebugInfoLogger.LogText("Aggiunta righe manganti alla tabella Sintesi", "OK");

            // Aggiorno le formule presenti nella parte destra della tabella del foglio "Sintesi"
            AggiornaFormuleInTabelleSintesi(context.InfoFileReport, context.Configurazione, context.UpdateReportsInput.Periodo);
            context.DebugInfoLogger.LogText("Aggiornamento delle formule nella tabella Sintesi", context.UpdateReportsInput.Periodo);
            #endregion

            // La lista delle tuple viene creata a partire dalle spese, questo potrebbe portare ad avere
            // combinazioni di (fornitore/reparto/tipologia di spese) presenti in Tabella sintesi
            // ma non presenti in "tupleSpesePerAnno" in quanto per queste combinazioni non sono presenti spese
            // Aggiungo quindi a "tupleSpesePerAnno" le tuple individuate in Tabella Sintesi ma mancanti alla lista
            AggiungiTupleSpese_PerRigheTabellaSintesi_SenzaSpese(tupleSpesePerAnno, righeTabellaSintesi);

            #region Operazione su tabella "Pianificazione con consumi"
            // Calcolo dei nuovi valori per la tabella applicando le spesa actual calcolate e riportando l'eventuale delta nei mesi futuri.
            DeterminaNuoviValoriPer_PianificazioneConConsumi(tupleSpesePerAnno, righeTabellaSintesi, context.UpdateReportsInput.Periodo);
            context.DebugInfoLogger.LogRigheTabellaSintesi_NewValues(righeTabellaSintesi);

            // Duplicazione dei dati da foglio "Sintesi" a "Pianificazione con consumi"
            CopiaValoriDa_Sintesi_in_PianificazioneConConsumi(context.InfoFileReport, context.Configurazione);
            context.DebugInfoLogger.LogText("Copia dei dati dal foglio Sintesi a Pianificazione con consumi", "OK");

            //Registra i valori calcolati per i seguenti fogli: "Pianificazione con consumi", "Spese actual", "Spese commitment", "Spese totali"
            Aggiorna_PianificazioneConConsumi(context.InfoFileReport, context.Configurazione, context.RepartiCensitiInReport, context.FornitoriCensitiInReport, context.RigheSpese, righeTabellaSintesi, out List<RigaLogElaborazioniTabellaSintesi> logOperazioniPianificazioneConConsumi);
            context.DebugInfoLogger.LogText("Aggiornamento dei valori nella tabella Pianificazione con consumi", "OK");
            context.DebugInfoLogger.LogOperazioniSuPianificazioneConConsumi(logOperazioniPianificazioneConConsumi);

            // Scrivo i cambiamenti applicati al fogli sintesi nelle due tabelle presenti nel foglio "Sintesi-Log modifiche"
            RegistraOperazioniSuPianificazioneConConsumi_SuFoglioLogModifiche(context.InfoFileReport, logOperazioniPianificazioneConConsumi, context.UpdateReportsInput.Periodo);
            context.DebugInfoLogger.LogText("Registrazione dei valori nelle tabelle LogModifica", "OK");
            #endregion

            #region Fogli spese
            // Aggiorna i fogli "Spese"
            Aggiorna_FogliSpese(context.InfoFileReport, context.Configurazione, tupleSpesePerAnno);
            #endregion

            #region Operazioni comuni a tutti i fogli
            // Sposto il marcatore della periodo nella nuova posizione su entrambi i fogli "Sintesi"
            AggiornaPosizioneMarcatorePeriodo(context.InfoFileReport, context.Configurazione, context.UpdateReportsInput.Periodo);
            context.DebugInfoLogger.LogText("Aggiornamento marcatore periodo sui mesi della tabella Sintesi", context.UpdateReportsInput.Periodo);


            // Setta il colore backgroud alle celle dei mesi passati
            SettaColoreBackgroundAlleCellePeriodoPassato(context.InfoFileReport, context.Configurazione, context.UpdateReportsInput.Periodo, righeTabellaSintesi.Count);
            context.DebugInfoLogger.LogText("Settaggio colore di background per le celle dei mesi nel passato", "OK");
            #endregion

            return null;
        }



        /// <summary>
        /// Determina il set minimo di tuple (Fornitore, Reparti, TipologiaDiSpesa che la tabella sintesi dovrebbe avere
        /// </summary>
        private List<TuplaSpesePerAnno> CreaSetMinimoTupleSpese(List<RigaTabellaAvanzamento> righeTabellaAvanzamento)
        {
            // Per ognuna delle due tipologie di spesa ("Lump sum" e "Ore") seleziono tutte le righe con un totale
            // di spesa diverso da zero
            var speseTabellaSintesi = new List<TuplaSpesePerAnno>();

            // tuple per tipologie di spesa "Lump sum"
            var tupleSpesa_LumpSum = righeTabellaAvanzamento.Where(_ => _.TotaleSpeseLumpSum != 0)
                        .Distinct()
                        .Select(_ => new TuplaSpesePerAnno(_.SiglaFornitore, _.Reparto, TipologieDiSpesa.LumpSum));
            speseTabellaSintesi.AddRange(tupleSpesa_LumpSum);

            // tuple per tipologie di spesa "Ore"
            var tupleSpesa_Ore = righeTabellaAvanzamento.Where(_ => _.TotaleSpeseAdOre_Euro != 0)
                        .Distinct()
                        .Select(_ => new TuplaSpesePerAnno(_.SiglaFornitore, _.Reparto, TipologieDiSpesa.AdOre));
            speseTabellaSintesi.AddRange(tupleSpesa_Ore);

            return speseTabellaSintesi.OrderBy(_ => _.Reparto)
                                    .ThenBy(_ => _.SiglaFornitore)
                                    .ThenBy(_ => _.TipologiaDiSpesa)
                                    .ToList();
        }

        private List<TuplaSpesePerAnno> CreaSetMinimoTupleSpese_V2(List<RigaSpese> righeSpesa)
        {
            var speseTabellaSintesi = new List<TuplaSpesePerAnno>();

            foreach (var rigaSpese in righeSpesa)
            {
                if (speseTabellaSintesi.Any(_ =>
                _.SiglaFornitore == rigaSpese.Fornitore.SiglaInReport &&
                _.Reparto == rigaSpese.NomeReparto &&
                _.TipologiaDiSpesa == rigaSpese.TipologiaDiSpesa))
                { continue; }

                speseTabellaSintesi.Add(new TuplaSpesePerAnno(rigaSpese.Fornitore.SiglaInReport, rigaSpese.NomeReparto, rigaSpese.TipologiaDiSpesa));
            }

            return speseTabellaSintesi.OrderBy(_ => _.Reparto)
                                    .ThenBy(_ => _.SiglaFornitore)
                                    .ThenBy(_ => _.TipologiaDiSpesa)
                                    .ToList();
        }


        /// <summary>
        /// Legge le righe attualmente presenti nella tabella sintesi (dati identificativi tupla, budget stansiato)
        /// </summary>
        private List<RigaTabellaSintesi> GetRigheTabellaSintesi(InfoFileReport infoFileReport, Configurazione configurazione, List<Reparto> repartiCensitiInReport, List<FornitoreCensito> fornitoriCensiti)
        {
            var worksheetName = infoFileReport.WorksheetName_Sintesi;  // Sintesi

            var righeAttuali = new List<RigaTabellaSintesi>();
            var ultimaRigaUsataNelFoglio = infoFileReport.EPPlusHelper.GetRowsLimit(worksheetName);
            var rigaCorrente = configurazione.Sintesi_Riga_PrimaConDati;

            while (rigaCorrente <= ultimaRigaUsataNelFoglio)
            {
                var siglaFornitore = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.Sintesi_ColonnaSiglaFornitore);
                var reparto = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.Sintesi_ColonnaReparto);
                var objectTipologiaSpesa = infoFileReport.EPPlusHelper.GetValue(worksheetName, rigaCorrente, configurazione.Sintesi_ColonnaTipologiaSpesa);

                // se tutte i valori sono nulli ho raggiunto la fine della tabella, mi fermo
                if (siglaFornitore == null && reparto == null && objectTipologiaSpesa == null)
                { break; }

                // se uno o più valori è nullo sollevo un'eccezione
                if (siglaFornitore == null || reparto == null || objectTipologiaSpesa == null)
                {
                    throw new ManagedException(
                    tipologiaErrore: TipologiaErrori.DatoMancante,
                    tipologiaCartella: TipologiaCartelle.ReportInput,
                    worksheetName: worksheetName,
                    rigaCella: rigaCorrente,
                    colonnaCella: siglaFornitore == null ? configurazione.Sintesi_ColonnaSiglaFornitore
                                : reparto == null ? configurazione.Sintesi_ColonnaReparto
                                : objectTipologiaSpesa == null ? configurazione.Sintesi_ColonnaTipologiaSpesa
                                : -1,
                    nomeDatoErrore: siglaFornitore == null ? NomiDatoErrore.SiglaFornitore
                                : reparto == null ? NomiDatoErrore.NomeReparto
                                : objectTipologiaSpesa == null ? NomiDatoErrore.TipologiaSpesa
                                : NomiDatoErrore.None,
                    dato: null,
                    percorsoFile: null
                    );
                }

                // Validazione dei valori letti
                if (!ValueHelper.CastTipologieDiSpesa(objectTipologiaSpesa, out TipologieDiSpesa? tipologieDiSpesa))
                {
                    // questo errore non dovrebbe potersi verificare in quanto i valori inseribili nelle celle sono vincolati da un menù a tendina
                    throw new ManagedException(
                          tipologiaErrore: TipologiaErrori.DatoNonValido,
                          tipologiaCartella: TipologiaCartelle.ReportInput,
                          worksheetName: worksheetName,
                          rigaCella: rigaCorrente,
                          colonnaCella: configurazione.Sintesi_ColonnaTipologiaSpesa,
                          nomeDatoErrore: NomiDatoErrore.TipologiaSpesa,
                          dato: objectTipologiaSpesa.ToString()
                          );
                }

                // verifico che la Sigla del fornitore sia valida
                if (!fornitoriCensiti.Any(_ => _.SiglaInReport.Equals(siglaFornitore, StringComparison.InvariantCultureIgnoreCase)))
                {
                    throw new ManagedException(
                          tipologiaErrore: TipologiaErrori.DatoNonValido,
                          tipologiaCartella: TipologiaCartelle.ReportInput,
                          worksheetName: worksheetName,
                          rigaCella: rigaCorrente,
                          colonnaCella: configurazione.Sintesi_ColonnaSiglaFornitore,
                          nomeDatoErrore: NomiDatoErrore.SiglaFornitore,
                          dato: siglaFornitore
                          );
                }

                // verifico che il Reparto sia un valore valido
                if (!repartiCensitiInReport.Any(_ => _.Nome.Equals(reparto, StringComparison.InvariantCultureIgnoreCase)))
                {
                    // questo errore non dovrebbe potersi verificare in quanto i valori inseribili nelle celle sono vincolati da un menù a tendina
                    throw new ManagedException(
                          tipologiaErrore: TipologiaErrori.DatoNonValido,
                          tipologiaCartella: TipologiaCartelle.ReportInput,
                          worksheetName: worksheetName,
                          rigaCella: rigaCorrente,
                          colonnaCella: configurazione.Sintesi_ColonnaReparto,
                          nomeDatoErrore: NomiDatoErrore.NomeReparto,
                          dato: reparto
                          );
                }

                var categoriaFornitori = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.Sintesi_ColonnaCategoriaFornitore);

                // Predisgongo una nuova riga
                var rigaAttuale = new RigaTabellaSintesi(siglaFornitore, reparto, tipologieDiSpesa.Value, categoriaFornitori);
                // Leggo i valori nei mesi
                for (var mese = 0; mese <= 11; mese++)
                {
                    var valoreAttualeMese = (double?)infoFileReport.EPPlusHelper.GetValue(worksheetName, rigaCorrente, configurazione.Sintesi_Colonna_PrimaRangeMesi + mese);
                    if (valoreAttualeMese.HasValue)
                    {
                        rigaAttuale.BudgetStanziato[mese] = valoreAttualeMese.Value;
                    }
                }
                righeAttuali.Add(rigaAttuale);

                // vado alla riga successiva
                rigaCorrente++;
            }

            return righeAttuali;
        }


        /// <summary>
        /// Individua la lista di righe mancanti tra quelle presenti e quelle richiesta 
        /// dal set minimo al fine da coprire tutte le combinazioni per le spese presenti
        /// </summary>
        private List<RigaTabellaSintesi> IndividuaRigheMancantiTabellaSintesi(List<RigaTabellaSintesi> righeTabellaSintesi, List<TuplaSpesePerAnno> tupleSpese, List<FornitoreCensito> fornitoriCensiti)
        {
            var righeTabellaSintesiDaAggiungere = new List<RigaTabellaSintesi>();
            foreach (var tuplaSpeseTabellaSintesi in tupleSpese)
            {
                if (!righeTabellaSintesi.Any(_ =>
                            _.SiglaFornitore.Equals(tuplaSpeseTabellaSintesi.SiglaFornitore, StringComparison.InvariantCultureIgnoreCase) &&
                            _.Reparto.Equals(tuplaSpeseTabellaSintesi.Reparto, StringComparison.InvariantCultureIgnoreCase) &&
                            _.TipologiaDiSpesa.Equals(tuplaSpeseTabellaSintesi.TipologiaDiSpesa)))
                {
                    // Agiungo le eventuali riche che risultano mancanti
                    var rigaDaAggiungere = new RigaTabellaSintesi(tuplaSpeseTabellaSintesi.SiglaFornitore,
                                                    tuplaSpeseTabellaSintesi.Reparto,
                                                    tuplaSpeseTabellaSintesi.TipologiaDiSpesa,
                                                    fornitoriCensiti.Single(f => f.SiglaInReport.Equals(tuplaSpeseTabellaSintesi.SiglaFornitore)).Categoria);

                    righeTabellaSintesiDaAggiungere.Add(rigaDaAggiungere);
                }
            }
            return righeTabellaSintesiDaAggiungere;
        }


        /// <summary>
        /// Aggiunge le righe mancanti in fondo alla tabella sintesi e alla delle righe presenti
        /// </summary>
        private void AggiungiRigheMancantiTabelleSintesi(InfoFileReport infoFileReport, Configurazione configurazione, List<RigaTabellaSintesi> righeTabellaSintesi, List<RigaTabellaSintesi> righeMancanti)
        {
            righeTabellaSintesi.AddRange(righeMancanti);

            var worksheetName = infoFileReport.WorksheetName_Sintesi;  // Sintesi
            var rigaDaAggiornare = infoFileReport.EPPlusHelper.GetFirstEmptyRow(worksheetName, configurazione.Sintesi_Riga_PrimaConDati, configurazione.Sintesi_ColonnaSiglaFornitore, configurazione.Sintesi_ColonnaReparto, configurazione.Sintesi_ColonnaTipologiaSpesa);
            foreach (var rigaTabellaSintesiDaAggiungere in righeMancanti)
            {
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaDaAggiornare, configurazione.Sintesi_ColonnaSiglaFornitore, rigaTabellaSintesiDaAggiungere.SiglaFornitore, true);
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaDaAggiornare, configurazione.Sintesi_ColonnaReparto, rigaTabellaSintesiDaAggiungere.Reparto, true);
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaDaAggiornare, configurazione.Sintesi_ColonnaTipologiaSpesa, rigaTabellaSintesiDaAggiungere.TipologiaDiSpesa.GetEnumDescription(), true);
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaDaAggiornare, configurazione.Sintesi_ColonnaCategoriaFornitore, rigaTabellaSintesiDaAggiungere.CategoriaFornitori, true);
                rigaDaAggiornare++;
            }
        }


        /// <summary>
        /// 
        /// </summary>
        private void AggiornaFormuleInTabelleSintesi(InfoFileReport infoFileReport, Configurazione configurazione, int periodo)
        {
            var worksheetName = infoFileReport.WorksheetName_Sintesi;  // Sintesi

            #region Aggiornamento delle formule nella prima riga
            // Range delle celle con i valori di spesa per i 12 mesi: L19:W19
            // quindi L=Gen, M=Feb, W=Dic
            // Esempio periodo Aprile ovvero 4
            // Formula ad oggi:     Somma da Gen a Mar  ovvero da 1 a 3         in termini di colonne Excel da L a O
            // Formula a  finire:   Somma da Apr a Dic  ovvero da 4 a 12        in termini di colonne Excel da P a W
            // Generalizzando per periodo <X>
            // Formula ad oggi:     Somma da Gen a Mar  ovvero da 1 a <X-1>     in termini di colonne Excel da L a <X>
            // Formula a  finire:   Somma da Apr a Dic  ovvero da <X> a 12      in termini di colonne Excel da <X> a W

            // corrisponde a <X-1> nell'esempio
            var colonnaFineRangeFormula_Ad_Oggi = ColumnIDS.L + periodo - 1 - 1; // -1 = offset rispetto a L
                                                                                 // corrisponde a <X> nell'esempio
            var colonnaInizioRangeFormula_A_Finire = ColumnIDS.L + periodo - 1; // -1 = offset rispetto a L

            // N.B. Le formule devono esssere create secondo il formato di Ecxel inglese            
            //Formula_Finire = "=IF($K19=\"{0}\",SUM({1}19:W19),0)"; // Somma valori da colonna <X> a W            
            //Formula_AdOggi = "=IF($K19=\"{0}\",SUM(L19:{1}19),0)"; // Somma valori da colonna L a <X-1>

            // formula a finire, "ORE"
            var formulaDaIncollare = string.Format(configurazione.Sintesi_Formula_Finire, TipologieDiSpesa.AdOre.GetEnumDescription(), colonnaInizioRangeFormula_A_Finire);
            infoFileReport.EPPlusHelper.SetFormula(worksheetName, configurazione.Sintesi_Riga_PrimaConDati, configurazione.Sintesi_Colonna_Formula_Finire_Ore, formulaDaIncollare);

            // formula a finire, "Lump sum"
            formulaDaIncollare = string.Format(configurazione.Sintesi_Formula_Finire, TipologieDiSpesa.LumpSum.GetEnumDescription(), colonnaInizioRangeFormula_A_Finire);
            infoFileReport.EPPlusHelper.SetFormula(worksheetName, configurazione.Sintesi_Riga_PrimaConDati, configurazione.Sintesi_Colonna_Formula_Finire_LumpSum, formulaDaIncollare);

            // formula ad oggi, "ORE"
            formulaDaIncollare = (periodo != 1)
                                ? string.Format(configurazione.Sintesi_Formula_AdOggi, TipologieDiSpesa.AdOre.GetEnumDescription(), colonnaFineRangeFormula_Ad_Oggi)
                                : "0"; // a gennaio il valore è zero
            infoFileReport.EPPlusHelper.SetFormula(worksheetName, configurazione.Sintesi_Riga_PrimaConDati, configurazione.Sintesi_Colonna_Formula_AdOggi_Ore, formulaDaIncollare);

            // formula ad oggi, "Lump sum"
            formulaDaIncollare = (periodo != 1)
                                ? string.Format(configurazione.Sintesi_Formula_AdOggi, TipologieDiSpesa.LumpSum.GetEnumDescription(), colonnaFineRangeFormula_Ad_Oggi)
                                : "0"; // a gennaio il valore è zero
            infoFileReport.EPPlusHelper.SetFormula(worksheetName, configurazione.Sintesi_Riga_PrimaConDati, configurazione.Sintesi_Colonna_Formula_AdOggi_LumpSum, formulaDaIncollare);
            #endregion

            var primaRigaInCuiCopiareLeFormule = configurazione.Sintesi_Riga_PrimaConDati + 1;
            var ultimaRigaInCuiCopiareLeFormule = infoFileReport.EPPlusHelper.GetLastUsedRowForColumn(worksheetName, primaRigaInCuiCopiareLeFormule, configurazione.Sintesi_Colonna_Formula_Finire_Ore);

            #region  Copia e incolla delle formule dalla prima riga a tutte le altre nella colonna
            for (var rigaDaAggiornare = primaRigaInCuiCopiareLeFormule; rigaDaAggiornare <= ultimaRigaInCuiCopiareLeFormule; rigaDaAggiornare++)
            {
                infoFileReport.EPPlusHelper.CopiaFormula(worksheetName, configurazione.Sintesi_Riga_PrimaConDati, configurazione.Sintesi_Colonna_Formula_Finire_Ore,
                                                                     rigaDaAggiornare, configurazione.Sintesi_Colonna_Formula_Finire_Ore);
                infoFileReport.EPPlusHelper.CopiaFormula(worksheetName, configurazione.Sintesi_Riga_PrimaConDati, configurazione.Sintesi_Colonna_Formula_Finire_LumpSum,
                                                                     rigaDaAggiornare, configurazione.Sintesi_Colonna_Formula_Finire_LumpSum);
                infoFileReport.EPPlusHelper.CopiaFormula(worksheetName, configurazione.Sintesi_Riga_PrimaConDati, configurazione.Sintesi_Colonna_Formula_AdOggi_Ore,
                                                                     rigaDaAggiornare, configurazione.Sintesi_Colonna_Formula_AdOggi_Ore);
                infoFileReport.EPPlusHelper.CopiaFormula(worksheetName, configurazione.Sintesi_Riga_PrimaConDati, configurazione.Sintesi_Colonna_Formula_AdOggi_LumpSum,
                                                                     rigaDaAggiornare, configurazione.Sintesi_Colonna_Formula_AdOggi_LumpSum);
            }
            #endregion
        }

        private void AggiungiTupleSpese_PerRigheTabellaSintesi_SenzaSpese(List<TuplaSpesePerAnno> tupleSpesePerAnno, List<RigaTabellaSintesi> righeTabellaSintesi)
        {
            foreach (var rigaTabellaSintesi in righeTabellaSintesi)
            {
                if (tupleSpesePerAnno.Any(_ =>
                _.SiglaFornitore == rigaTabellaSintesi.SiglaFornitore &&
                _.Reparto == rigaTabellaSintesi.Reparto &&
                _.TipologiaDiSpesa == rigaTabellaSintesi.TipologiaDiSpesa))
                { continue; }

                tupleSpesePerAnno.Add(new TuplaSpesePerAnno(rigaTabellaSintesi.SiglaFornitore, rigaTabellaSintesi.Reparto, rigaTabellaSintesi.TipologiaDiSpesa));
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void AggiornaPosizioneMarcatorePeriodo(InfoFileReport infoFileReport, Configurazione configurazione, int periodo)
        {
            // Lista dei fogli su cui aggiornare il Marcatore del mese del periodo
            var worksheetNameTo_List = new List<string>
                {
                infoFileReport.WorksheetName_Sintesi,
                infoFileReport.WorksheetName_PianificazioneConConsumi,
                infoFileReport.WorksheetName_SpeseActual,
                infoFileReport.WorksheetName_SpeseComittment,
                infoFileReport.WorksheetName_SpeseTotali
                };

            foreach (var worksheetNameTo in worksheetNameTo_List)
            {
                var primaColonnaMesi = (worksheetNameTo.Equals(infoFileReport.WorksheetName_Sintesi) || worksheetNameTo.Equals(infoFileReport.WorksheetName_PianificazioneConConsumi))
                                        ? configurazione.Sintesi_Colonna_PrimaRangeMesi
                                        : configurazione.Spese_Colonna_PrimaRangeMesi;
                var rigaIntestazioneMesi = (worksheetNameTo.Equals(infoFileReport.WorksheetName_Sintesi) || worksheetNameTo.Equals(infoFileReport.WorksheetName_PianificazioneConConsumi))
                                        ? configurazione.Sintesi_Riga_IntestazioneMesi
                                        : configurazione.Spese_Riga_IntestazioneMesi;
                _aggiornaPosizioneMarcatorePeriodoSuFoglio(infoFileReport, configurazione, periodo, worksheetNameTo, rigaIntestazioneMesi, primaColonnaMesi);
            }
        }
        private void _aggiornaPosizioneMarcatorePeriodoSuFoglio(InfoFileReport infoFileReport, Configurazione configurazione, int periodo, string worksheetName, int rigaIntestazioneMesi, int primaColonnaMesi)
        {
            var colonnaDaEvidenziarePerPeriodo = (periodo == 1)
                            ? primaColonnaMesi + 11
                            : primaColonnaMesi + periodo - 2;

            // azzera i bordi correnti
            infoFileReport.EPPlusHelper.SetThinBorderOnRange(worksheetName, rigaIntestazioneMesi, primaColonnaMesi, configurazione.Sintesi_Riga_IntestazioneMesi, primaColonnaMesi + 11);

            // marca il bordo destro della cella del mese del periodo indicato
            infoFileReport.EPPlusHelper.SetThickRightBorder(worksheetName, rigaIntestazioneMesi, colonnaDaEvidenziarePerPeriodo);
        }


        /// <summary>
        /// Copia il contenuto del foglio "Sintesi" nel nuovo foglio "Sintesi ???"
        /// </summary>
        private void CopiaValoriDa_Sintesi_in_PianificazioneConConsumi(InfoFileReport infoFileReport, Configurazione configurazione)
        {
            _svuotaTabella_PianificazioneConConsumi(infoFileReport, configurazione);
            _copiaValoriDa_Sintesi_in_PianificazioneConConsumi(infoFileReport, configurazione);
        }
        private void _svuotaTabella_PianificazioneConConsumi(InfoFileReport infoFileReport, Configurazione configurazione)
        {
            //// Lista dei fogli da svuotare_FogliGemelli
            //var worksheetNameTo_List = new List<string>
            //    {
            //    infoFileReport.WorksheetName_PianificazioneConConsumi,
            //    infoFileReport.WorksheetName_SpeseActual,
            //    infoFileReport.WorksheetName_SpeseComittment,
            //    infoFileReport.WorksheetName_SpeseTotali
            //    };

            var worksheetName = infoFileReport.WorksheetName_PianificazioneConConsumi;

            //foreach (var worksheetName in worksheetNameTo_List)
            //{
            // svuoto il contenuto del foglio
            var ultimaRigaUsataNelFoglio = infoFileReport.EPPlusHelper.GetRowsLimit(worksheetName);
            infoFileReport.EPPlusHelper.CleanCellsContent(worksheetName, configurazione.Sintesi_Riga_PrimaConDati, 1, ultimaRigaUsataNelFoglio, configurazione.Sintesi_ColonnaOverBudget);
            //}
        }
        private void _copiaValoriDa_Sintesi_in_PianificazioneConConsumi(InfoFileReport infoFileReport, Configurazione configurazione)
        {
            // copio la i dati dal foglio "Sintesi"
            var worksheetNameFrom = infoFileReport.WorksheetName_Sintesi;
            var worksheetNameTo = infoFileReport.WorksheetName_PianificazioneConConsumi;


            //// incollo i valori del foglio "Sintesi" nei seguenti fogli
            //var worksheetNameTo_List = new List<string>
            //    {
            //    infoFileReport.WorksheetName_PianificazioneConConsumi,
            //    infoFileReport.WorksheetName_SpeseActual,
            //    infoFileReport.WorksheetName_SpeseComittment,
            //    infoFileReport.WorksheetName_SpeseTotali
            //    };

            var ultimaRigaUsataNelFoglio = infoFileReport.EPPlusHelper.GetRowsLimit(worksheetNameFrom);
            //foreach (var worksheetNameTo in worksheetNameTo_List)
            //{
            // duplicazione valori da "Sintesi" a <worksheetNameTo>
            var rigaCorrente = configurazione.Sintesi_Riga_IntestazioneMesi;
            while (rigaCorrente <= ultimaRigaUsataNelFoglio)
            {

                for (var colonnaCorrente = 1; colonnaCorrente <= configurazione.Sintesi_Colonna_UltimaRangeMesi; colonnaCorrente++)
                {
                    var isFormula = infoFileReport.EPPlusHelper.IsFormula(worksheetNameFrom, rigaCorrente, colonnaCorrente);
                    if (isFormula)
                    {
                        var formula = infoFileReport.EPPlusHelper.GetFormula(worksheetNameFrom, rigaCorrente, colonnaCorrente);
                        infoFileReport.EPPlusHelper.SetFormula(worksheetNameTo, rigaCorrente, colonnaCorrente, formula);
                    }
                    else
                    {
                        var valore = infoFileReport.EPPlusHelper.GetValue(worksheetNameFrom, rigaCorrente, colonnaCorrente);
                        infoFileReport.EPPlusHelper.SetValue(worksheetNameTo, rigaCorrente, colonnaCorrente, valore);
                    }
                }

                // vado alla riga successiva
                rigaCorrente++;
            }
            // }
        }

        //private void _copiaValori_DaFoglioSintesi_a_FogliGemelli(InfoFileReport infoFileReport, Configurazione configurazione)
        //{
        //    // copio la i dati dal foglio "Sintesi"
        //    var worksheetNameFrom = infoFileReport.WorksheetName_Sintesi;

        //    // incollo i valori del foglio "Sintesi" nei seguenti fogli
        //    var worksheetNameTo_List = new List<string>
        //        {
        //        infoFileReport.WorksheetName_PianificazioneConConsumi,
        //        infoFileReport.WorksheetName_SpeseActual,
        //        infoFileReport.WorksheetName_SpeseComittment,
        //        infoFileReport.WorksheetName_SpeseTotali
        //        };

        //    var ultimaRigaUsataNelFoglio = infoFileReport.EPPlusHelper.GetRowsLimit(worksheetNameFrom);
        //    foreach (var worksheetNameTo in worksheetNameTo_List)
        //    {
        //        // aggiornamento delle celle contenenti l'intestazione dei mesi. Per gestire il cambio di anno
        //        for (var mese = 0; mese <= 11; mese++)
        //        {
        //            var colonnaCorrente = configurazione.Sintesi_Colonna_PrimaRangeMesi + mese;
        //            var valore = infoFileReport.EPPlusHelper.GetValue(worksheetNameFrom, configurazione.Sintesi_Riga_IntestazioneMesi, colonnaCorrente);
        //            infoFileReport.EPPlusHelper.SetValue(worksheetNameTo, configurazione.Sintesi_Riga_IntestazioneMesi, colonnaCorrente, valore);
        //        }

        //        // duplicazione valori da "Sintesi" a <worksheetNameTo>
        //        var rigaCorrente = configurazione.Sintesi_Riga_PrimaConDati;
        //        while (rigaCorrente <= ultimaRigaUsataNelFoglio)
        //        {

        //            for (var colonnaCorrente = 1; colonnaCorrente <= configurazione.Sintesi_Colonna_PrimaRangeMesi - 1; colonnaCorrente++)
        //            {
        //                var isFormula = infoFileReport.EPPlusHelper.IsFormula(worksheetNameFrom, rigaCorrente, colonnaCorrente);
        //                if (isFormula)
        //                {
        //                    var formula = infoFileReport.EPPlusHelper.GetFormula(worksheetNameFrom, rigaCorrente, colonnaCorrente);
        //                    infoFileReport.EPPlusHelper.SetFormula(worksheetNameTo, rigaCorrente, colonnaCorrente, formula);
        //                }
        //                else
        //                {
        //                    var valore = infoFileReport.EPPlusHelper.GetValue(worksheetNameFrom, rigaCorrente, colonnaCorrente);
        //                    infoFileReport.EPPlusHelper.SetValue(worksheetNameTo, rigaCorrente, colonnaCorrente, valore);
        //                }
        //            }

        //            // vado alla riga successiva
        //            rigaCorrente++;
        //        }
        //    }
        //}


        /// <summary>
        /// Valorizza le collections ActualValues e CommittedValues con i mesi relativi ad ogni mese
        /// </summary>
        private void AssegnaTotaliSpeseMensiliAlleTuple(List<TuplaSpesePerAnno> tupleSpesePerAnno, List<RigaSpese> righeSpesa, int annoCorrente)
        {
            // Scorro tutte le righe di spese escludendo quelle a zero, in quanto per queste
            // ultime potremmo anche non avere una corrispondente tupla nella tabella "Sintesi"
            // foreach (var rigaSpese in righeSpesa.Where(_ => Math.Round(_.Spesa, 2) != 0))
            foreach (var rigaSpese in righeSpesa.Where(_ => _.Spesa != 0))
            {
                //var tuplaSpeseDaAggiornare = tupleSpesePerAnno.Where(_ =>
                //            _.SiglaFornitore.Equals(rigaSpese.Fornitore.SiglaInReport, StringComparison.InvariantCultureIgnoreCase) &&
                //            _.Reparto.Equals(rigaSpese.NomeReparto, StringComparison.InvariantCultureIgnoreCase) &&
                //            _.TipologiaDiSpesa.Equals(rigaSpese.TipologiaDiSpesa))
                //            .ToList();
                var tupleSpesePerAnnoDaAggiornare = tupleSpesePerAnno.SingleOrDefault(_ =>
                        _.SiglaFornitore.Equals(rigaSpese.Fornitore.SiglaInReport, StringComparison.InvariantCultureIgnoreCase) &&
                        _.Reparto.Equals(rigaSpese.NomeReparto, StringComparison.InvariantCultureIgnoreCase) &&
                        _.TipologiaDiSpesa.Equals(rigaSpese.TipologiaDiSpesa));
                if (tupleSpesePerAnnoDaAggiornare == null)
                {
                    // questa situazione, rara, ma possibile, si verifica quando ho delle spese per la combinazione (fornitore/reparto/tipologia di spese)
                    // diverse da zero ma la cui somma da zero.
                    // Questo comporta che la combinazione (fornitore/reparto/tipologia di spese)
                    // ma
                    continue;
                }


                // double valoreDaAssegnareAdOgniMeseInIntervallo = 0;

                // determino il numero di mesi su cui splittare la spesa (sulla base della data di inizio e fine)
                var numeroMesiInIntervalloDellaSpesa = rigaSpese.NumeroMesiSplitSpesaInTabellaSintesi;
                //if (numeroMesiInIntervalloDellaSpesa > 0)
                //{
                // in base al Tipologia di spesa prendo il valore di Spesa o delle Ore e
                // divido il valore per il numero di mesi su cui va spalmata la spesa
                var spesa = (rigaSpese.TipologiaDiSpesa == TipologieDiSpesa.LumpSum)
                            ? rigaSpese.Spesa
                            : rigaSpese.Ore.Value;
                var valoreDaAssegnareAdOgniMeseInIntervallo = Math.Round(spesa / numeroMesiInIntervalloDellaSpesa, Numbers.NumeroDecimaliImportiSpese);
                //  }

                //foreach (var rigaDaAggiornare in tuplaSpeseDaAggiornare)
                //{
                // Scorro di un mese alla volta il periodo su cui si estende la spesa per allocare ogni porzione di spesa nel mese corretto
                var dataInizioMeseCorrente = new DateTime(rigaSpese.DataInizioTabellaSintesi.Year, rigaSpese.DataInizioTabellaSintesi.Month, 1);
                while (dataInizioMeseCorrente < rigaSpese.DataFineTabellaSintesi)
                {
                    switch (rigaSpese.StatusSpesa)
                    {
                        // Al momento lascio separate le logiche di gestione in base allo StatusSpesa anche se identiche
                        // Logiche per spese di tipo Actual
                        case StatusSpesa.Actual:
                            if (dataInizioMeseCorrente.Year > annoCorrente)
                            {
                                // le porzioni di spesa che cadono nell'anno successivo vengono ignorate
                            }
                            else if (dataInizioMeseCorrente.Year < annoCorrente)
                            {
                                // le porzioni spese che cadono nell'anno precedente a quello di riferimento
                                // vengono tutte caricate sul mese di gennaio
                                tupleSpesePerAnnoDaAggiornare.SpeseActualPerMese[0] += valoreDaAssegnareAdOgniMeseInIntervallo;
                            }
                            else
                            {
                                // queste porzioni di spesa cadono nell'anno corrente:
                                // il valore da assegnare viene aggiunto al corrispondente mese di riferimento
                                tupleSpesePerAnnoDaAggiornare.SpeseActualPerMese[dataInizioMeseCorrente.Month - 1] += valoreDaAssegnareAdOgniMeseInIntervallo;
                            }
                            break;

                        // logiche per spese di tipo Commitment
                        case StatusSpesa.Commitment:
                            if (dataInizioMeseCorrente.Year > annoCorrente)
                            {
                                // le porzioni di spesa che cadono nell'anno successivo vengono ignorate
                            }
                            else if (dataInizioMeseCorrente.Year < annoCorrente)
                            {
                                // le porzioni spese che cadono nell'anno precedente a quello di riferimento
                                // vengono tutte caricate sul mese di gennaio
                                tupleSpesePerAnnoDaAggiornare.SpeseCommittedPerMese[0] += valoreDaAssegnareAdOgniMeseInIntervallo;
                            }
                            else
                            {
                                // queste porzioni di spesa cadono nell'anno corrente:
                                // il valore da assegnare viene aggiunto al corrispondente mese di riferimento
                                tupleSpesePerAnnoDaAggiornare.SpeseCommittedPerMese[dataInizioMeseCorrente.Month - 1] += valoreDaAssegnareAdOgniMeseInIntervallo;
                            }
                            break;
                    }

                    // avanzo al mese succesivo
                    dataInizioMeseCorrente = dataInizioMeseCorrente.AddMonths(1);
                }
                //}
            }
        }


        //todo: rendere private ma testabili
        /// <summary>
        /// Calcola i nuovi valori per la tabella sintesi sulla base delle spese actual processate
        /// </summary>
        public void DeterminaNuoviValoriPer_PianificazioneConConsumi(List<TuplaSpesePerAnno> tupleSpesePerAnno, List<RigaTabellaSintesi> righeTabellaSintesi, int periodo)
        {
            foreach (var tuplaSpese in tupleSpesePerAnno)
            {
                // Seleziono le righe esistenti in tabella per la tupla corrente
                var righeTabellaSintesiPerTupla = righeTabellaSintesi.Where(_ =>
                            _.SiglaFornitore.Equals(tuplaSpese.SiglaFornitore, StringComparison.InvariantCultureIgnoreCase) &&
                            _.Reparto.Equals(tuplaSpese.Reparto, StringComparison.InvariantCultureIgnoreCase) &&
                            _.TipologiaDiSpesa.Equals(tuplaSpese.TipologiaDiSpesa))
                            .ToList();

                // todo: questo if potrebbe essere superfluo
                if (!righeTabellaSintesiPerTupla.Any()) { continue; }

                _inizializza_ValorePianificazioneConConsumi_ConBudgetStanziato(righeTabellaSintesiPerTupla);

                _applicaValoriSpeseRealiComeValorePianificazioneConConsumi_MesiPassati(tuplaSpese, righeTabellaSintesiPerTupla, periodo, out double deltaAccumulato);

                _aggiornaValorePianificazioneConConsumi_MesiFuturi(tuplaSpese, righeTabellaSintesiPerTupla, periodo, deltaAccumulato);

                //_applicaSpeseRealiSeOltreBudgetComeValorePianificazioneConConsumi_MesiFuturi(tuplaSpese, righeTabellaSintesiPerTupla, periodo, deltaSpese_MesiPassati, out double deltaSpese_Totale);

                //_distribuisti_DeltaAccumulato(tuplaSpese, righeTabellaSintesiPerTupla, periodo, deltaSpese_Totale);

            }
        }
        private void _inizializza_ValorePianificazioneConConsumi_ConBudgetStanziato(List<RigaTabellaSintesi> righeTabellaSintesi/*, List<RigaSpese> righeSpese*/)
        {
            // Inizializzo tutti i valori NewValues con BudgetStanziato per poi successivamente ritoccarli in base ai valori di spesa
            foreach (var rigaAttualePerTuplaSpese in righeTabellaSintesi)
            {
                for (var mese = 0; mese <= 11; mese++)
                {
                    rigaAttualePerTuplaSpese.ValorePianificazioneConConsumi[mese] = rigaAttualePerTuplaSpese.BudgetStanziato[mese];
                }
            }
        }

        /// <summary>
        /// Per i mesi passati (prima del periodo), vengono aggiornati i nuovi valori facendoli corrispondere 
        /// alle spese sostenute. Viene calcolato il delta delle spese
        /// </summary>
        private void _applicaValoriSpeseRealiComeValorePianificazioneConConsumi_MesiPassati(TuplaSpesePerAnno tuplaSpese, List<RigaTabellaSintesi> righeAttualiPerTupla, int periodo, out double deltaAccumulato)
        {
            deltaAccumulato = 0;

            // Scorro i mesi passati (precedenti al perido) per determinare il delta tra pianificato ed actual
            for (var mese = 0; mese <= periodo - 2; mese++)
            {
                var budget_MeseCorrente = righeAttualiPerTupla.Sum(_ => _.BudgetStanziato[mese]);
                var spese_MeseCorrente = tuplaSpese.SpeseActualPerMese[mese] + tuplaSpese.SpeseCommittedPerMese[mese];
                var delta_MeseCorrente = spese_MeseCorrente - budget_MeseCorrente;

                // Accumulo il delta totale di tutti i mesi passati
                deltaAccumulato += delta_MeseCorrente;

                // Questo test con il round è necessario in quanto il deltaDiSpesa_MeseCorrente potrebbe essere molto vicino
                // a zero ma avere qualche frazione di decimale. Questo è dovuto al modo con cui le spese vengono splittate tra i mesi
                if (Math.Round(delta_MeseCorrente, Numbers.NumeroDecimaliImportiSpese) != 0)
                {
                    // Il delta del mese corrente è diverso da zero: applico le correzioni ai valori attuali
                    var listaValoriAttuali = righeAttualiPerTupla.Select(_ => _.BudgetStanziato[mese]).ToList();
                    var nuoviValori = (delta_MeseCorrente > 0)
                        // Delta positivo: le spese actual sono SUPERIORI a quelle previste devo AGGIUNGERE il delta ai valori current
                        ? ValueHelper.AddDeltaToValues_PriorityForNumberGreaterThanZero(listaValoriAttuali, delta_MeseCorrente)
                        // Delta negativo: le spese actual sono INFERIORI di quelle previste devo SOTTRARRE il delta ai valori current
                        : ValueHelper.RemoveDeltaFromValues(listaValoriAttuali, -delta_MeseCorrente, true, out _);
                    _applicaNuovi_ValorePianificazioneConConsumi_AlMese(righeAttualiPerTupla, mese, nuoviValori);
                }
            }
        }


        private void _aggiornaValorePianificazioneConConsumi_MesiFuturi(TuplaSpesePerAnno tuplaSpese, List<RigaTabellaSintesi> righeAttualiPerTupla, int periodo, double deltaAccumulato)
        {
            for (var mese = periodo - 1; mese <= 11; mese++)
            {
                #region Calcolo del delta per il mese corrente
                var budget_MeseCorrente = righeAttualiPerTupla.Sum(_ => _.BudgetStanziato[mese]);
                if (Math.Round(deltaAccumulato, Numbers.NumeroDecimaliImportiSpese) < 0)
                {
                    // Aggiungo il delta risparmiato nei mesi passati al budget del mese
                    budget_MeseCorrente += -deltaAccumulato;
                }
                var spese_MeseCorrente = tuplaSpese.SpeseActualPerMese[mese] + tuplaSpese.SpeseCommittedPerMese[mese];
                var delta_MeseCorrente = spese_MeseCorrente - budget_MeseCorrente;
                #endregion

                #region Caso delta_MeseCorrente positivo 
                if (Math.Round(delta_MeseCorrente, Numbers.NumeroDecimaliImportiSpese) > 0)
                {
                    // Delta positivo: le spese sono SUPERIORI al previsto.
                    // Devo AGGIUNGERE il delta ai valori attuali
                    var listaValoriAttuali = righeAttualiPerTupla.Select(_ => _.BudgetStanziato[mese]).ToList();
                    var nuoviValori = ValueHelper.AddDeltaToValues_PriorityForNumberGreaterThanZero(listaValoriAttuali, delta_MeseCorrente);
                    _applicaNuovi_ValorePianificazioneConConsumi_AlMese(righeAttualiPerTupla, mese, nuoviValori);

                    // Riporto il conteggio totale per il delta dei mesi futuri
                    deltaAccumulato += delta_MeseCorrente;
                }
                #endregion

                #region Caso delta_MeseCorrente negativo o uguale a zero 
                if (Math.Round(delta_MeseCorrente, Numbers.NumeroDecimaliImportiSpese) <= 0)
                {
                    // Il deltaDiSpesa_MeseCorrente è negativo o uguale a zero.
                    // In questo mese si è speso QUANTO previsto o MENO
                    // posso compensare un eventuale deltaAccumulato positivo
                    if (Math.Round(deltaAccumulato, Numbers.NumeroDecimaliImportiSpese) > 0)
                    {
                        // Ho un delta accoumulato maggiore di zero, cerco di compensarlo almeno inparte

                        // Compenso il valore più piccolo tra deltaDiSpesa_MeseCorrente e deltaAccumulato
                        var deltaDaCompensare = (deltaAccumulato > -delta_MeseCorrente)
                                                ? -delta_MeseCorrente
                                                : deltaAccumulato;
                        var listaValoriAttuali = righeAttualiPerTupla.Select(_ => _.ValorePianificazioneConConsumi[mese]).ToList();
                        var nuoviValori = ValueHelper.RemoveDeltaFromValues(listaValoriAttuali, deltaDaCompensare, false, out _ /*double remainingDelta*/);
                        _applicaNuovi_ValorePianificazioneConConsumi_AlMese(righeAttualiPerTupla, mese, nuoviValori);

                        deltaAccumulato -= deltaDaCompensare;
                    }
                    else if (Math.Round(deltaAccumulato, Numbers.NumeroDecimaliImportiSpese) < 0)
                    {
                        // Ho un delta accoumulato minore di zero, lo aggiungo al budget del mese corrente
                        var listaValoriAttuali = righeAttualiPerTupla.Select(_ => _.ValorePianificazioneConConsumi[mese]).ToList();
                        var nuoviValori = ValueHelper.AddDeltaToValues_PriorityForNumberGreaterThanZero(listaValoriAttuali, -deltaAccumulato);
                        _applicaNuovi_ValorePianificazioneConConsumi_AlMese(righeAttualiPerTupla, mese, nuoviValori);
                        deltaAccumulato = 0;
                    }
                }
                #endregion
            }

            // Verifico che il deltaSpese_Totale si stato effettivamente compensanto interamente
            if (Math.Round(deltaAccumulato, Numbers.NumeroDecimaliImportiSpese) > 0)
            {
                // deltaSpese_Totale accumulato è ancora maggiore di zero
                // Non c'è stato abbastanza spazio nei mesi futuri per assorbirlo, lo applico come extra budget
                var extraBudgetEquidistribuitoPerRiga = Math.Round(deltaAccumulato / righeAttualiPerTupla.Count, Numbers.NumeroDecimaliImportiSpese);

                for (var j = 0; j < righeAttualiPerTupla.Count; j++)
                {
                    //righeAttualiPerTupla[j].ExtraBudget = extraBudgetEquidistribuitoPerRiga;
                    var extraBudgetCalcolatoPerRiga = righeAttualiPerTupla[j].DeltaApplicato.Sum();
                    righeAttualiPerTupla[j].ExtraBudget = extraBudgetCalcolatoPerRiga;
                }
            }
        }



        /// <summary>
        /// Per i mesi futuri si tenta di conservare come valori NewValues i valori del BudgetStanziato  
        /// </summary>
        private void _applicaSpeseRealiSeOltreBudgetComeValorePianificazioneConConsumi_MesiFuturi(TuplaSpesePerAnno tuplaSpese, List<RigaTabellaSintesi> righeAttualiPerTupla, int periodo, double deltaSpese_MesiPassati, out double deltaSpese_Totale)
        {
            deltaSpese_Totale = deltaSpese_MesiPassati;

            // Osservo se nel mesi futuri ci sono spese più alte del budget previsto
            for (var mese = periodo - 1; mese <= 11; mese++)
            {
                var budget_MeseCorrente = righeAttualiPerTupla.Sum(_ => _.BudgetStanziato[mese]);
                var spese_MeseCorrente = tuplaSpese.SpeseActualPerMese[mese] + tuplaSpese.SpeseCommittedPerMese[mese];
                var delta_MeseCorrente = spese_MeseCorrente - budget_MeseCorrente;

                if (Math.Round(delta_MeseCorrente, Numbers.NumeroDecimaliImportiSpese) > 0)
                {
                    // Delta positivo: le spese sono SUPERIORI a quanto previste.
                    // Devo AGGIUNGERE il delta ai valori current
                    var listaValoriAttuali = righeAttualiPerTupla.Select(_ => _.BudgetStanziato[mese]).ToList();
                    var nuoviValori = ValueHelper.AddDeltaToValues_PriorityForNumberGreaterThanZero(listaValoriAttuali, delta_MeseCorrente);
                    _applicaNuovi_ValorePianificazioneConConsumi_AlMese(righeAttualiPerTupla, mese, nuoviValori);

                    // Riporto il conteggio totale per il delta dei mesi futuri
                    deltaSpese_Totale += delta_MeseCorrente;
                }
            }
        }



        private void _distribuisti_DeltaAccumulato(TuplaSpesePerAnno tuplaSpese, List<RigaTabellaSintesi> righeAttualiPerTupla, int periodo, double deltaSpese_Totale)
        {
            if (Math.Round(deltaSpese_Totale, Numbers.NumeroDecimaliImportiSpese) < 0)
            {
                // Delta negativo: in totale ho speso meno del previsto, aggiungo quanto avanzato al primo mese futuro
                var mese = periodo - 1;
                var listaValoriAttuali = righeAttualiPerTupla.Select(_ => _.ValorePianificazioneConConsumi[mese]).ToList();
                var nuoviValori = ValueHelper.AddDeltaToValues_PriorityForNumberGreaterThanZero(listaValoriAttuali, -deltaSpese_Totale);
                _applicaNuovi_ValorePianificazioneConConsumi_AlMese(righeAttualiPerTupla, mese, nuoviValori);

                deltaSpese_Totale = 0;

                return;
            }
            if (Math.Round(deltaSpese_Totale, Numbers.NumeroDecimaliImportiSpese) > 0)
            {
                // Delta positivo: in totale ho speso più del previsto, cerco di recuperare dai mesi futuri
                for (var mese = periodo - 1; mese <= 11; mese++)
                {
                    var spese_MeseCorrente = tuplaSpese.SpeseActualPerMese[mese] + tuplaSpese.SpeseCommittedPerMese[mese];
                    var budget_StanziatoMeseCorrente = righeAttualiPerTupla.Sum(_ => _.BudgetStanziato[mese]);
                    var deltaDiSpesa_MeseCorrente = spese_MeseCorrente - budget_StanziatoMeseCorrente;

                    if (Math.Round(deltaDiSpesa_MeseCorrente, Numbers.NumeroDecimaliImportiSpese) < 0)
                    {
                        // Il deltaDiSpesa_MeseCorrente è negativo: in questo mese si è speso MENO del previsto
                        // posso compenzare parte del deltaSpese_Totale

                        if (Math.Round(deltaSpese_Totale, Numbers.NumeroDecimaliImportiSpese) > 0)
                        {
                            // Compenso il valore più piccolo tra deltaDiSpesa_MeseCorrente e deltaAccumulato
                            var deltaDaCompensare = (deltaSpese_Totale > -deltaDiSpesa_MeseCorrente)
                                                    ? -deltaDiSpesa_MeseCorrente
                                                    : deltaSpese_Totale;
                            var listaValoriAttuali = righeAttualiPerTupla.Select(_ => _.ValorePianificazioneConConsumi[mese]).ToList();
                            var nuoviValori = ValueHelper.RemoveDeltaFromValues(listaValoriAttuali, deltaDaCompensare, false, out _ /*double remainingDelta*/);
                            _applicaNuovi_ValorePianificazioneConConsumi_AlMese(righeAttualiPerTupla, mese, nuoviValori);

                            deltaSpese_Totale -= deltaDaCompensare;
                        }
                    }
                }

                // Verifico che il deltaSpese_Totale si stato effettivamente compensanto interamente
                if (Math.Round(deltaSpese_Totale, Numbers.NumeroDecimaliImportiSpese) > 0)
                {
                    // deltaSpese_Totale accumulato è ancora maggiore di zero
                    // Non c'è stato abbastanza spazio nei mesi futuri per assorbirlo, lo applico come extra budget

                    var extraBudgetEquidistribuitoPerRiga = Math.Round(deltaSpese_Totale / righeAttualiPerTupla.Count, Numbers.NumeroDecimaliImportiSpese);


                    // codice per controlli incorciati
                    //var totaleDeltaApplicato = righeAttualiPerTupla.Sum(_ => _.DeltaApplicato.Sum());                
                    //if (deltaSpese_Totale != totaleDeltaApplicato)
                    //{
                    //    var s = "Strano!";
                    //}

                    for (var j = 0; j < righeAttualiPerTupla.Count; j++)
                    {
                        //righeAttualiPerTupla[j].ExtraBudget = extraBudgetEquidistribuitoPerRiga;

                        var extraBudgetCalcolatoPerRiga = righeAttualiPerTupla[j].DeltaApplicato.Sum();
                        righeAttualiPerTupla[j].ExtraBudget = extraBudgetCalcolatoPerRiga;

                    }
                }

            }



            return;
            for (var mese = periodo - 1; mese <= 11; mese++)
            {
                var speseMeseCorrente = tuplaSpese.SpeseActualPerMese[mese] + tuplaSpese.SpeseCommittedPerMese[mese];
                var budgetStanziatoMeseCorrente = righeAttualiPerTupla.Sum(_ => _.BudgetStanziato[mese]);

                // Calcolo, per il mese corrente, il delta tra i valori correnti nelle celle e quello derivato dalle spese
                var deltaDiSpesa_MeseCorrente = speseMeseCorrente - budgetStanziatoMeseCorrente;

                if (Math.Round(deltaDiSpesa_MeseCorrente, Numbers.NumeroDecimaliImportiSpese) > 0)
                {
                    // In questo mese si è speso PIU' del previsto:
                    // non posso compensare un eventuale deltaSpese_Totale positivo
                    deltaSpese_Totale += deltaDiSpesa_MeseCorrente;
                }

                if (Math.Round(deltaDiSpesa_MeseCorrente, Numbers.NumeroDecimaliImportiSpese) <= 0)
                {
                    // Il deltaDiSpesa_MeseCorrente è negativo:
                    // C'è spazio per recuperare parte del deltaAccumulato
                    // In questo mese si è speso MENO o quanto previsto:
                    // posso aggiungere un eventuale deltaSpese_Totale positivo

                    if (Math.Round(deltaSpese_Totale, Numbers.NumeroDecimaliImportiSpese) > 0)
                    {
                        // il deltaSpese_Totale di spesa è positivo (quindi maggiore del budget previsto)
                        // devo sottrarlo dal budget dei mesi futuri

                        // Compenso il valore più piccolo tra deltaDiSpesa_MeseCorrente e deltaAccumulato
                        var deltaDaCompensare = (deltaSpese_Totale > -deltaDiSpesa_MeseCorrente)
                                                ? -deltaDiSpesa_MeseCorrente
                                                : deltaSpese_Totale;
                        var listaValoriAttuali = righeAttualiPerTupla.Select(_ => _.ValorePianificazioneConConsumi[mese]).ToList();
                        var nuoviValori = ValueHelper.RemoveDeltaFromValues(listaValoriAttuali, deltaDaCompensare, false, out _ /*double remainingDelta*/);
                        _applicaNuovi_ValorePianificazioneConConsumi_AlMese(righeAttualiPerTupla, mese, nuoviValori);

                        deltaSpese_Totale -= deltaDaCompensare;
                    }
                }

                // Quando ho un deltaSpese_Totale negativo lo riverso nel primo mese succcessivo 
                if (Math.Round(deltaSpese_Totale, Numbers.NumeroDecimaliImportiSpese) < 0)
                {
                    // Caso deltaSpese_Totale negativo: le spese actual sono inferiori a quelle previste
                    // aggiungo il delta alle righe del primo mese futuro

                    var listaValoriAttuali = righeAttualiPerTupla.Select(_ => _.BudgetStanziato[mese]).ToList();
                    var nuoviValori = ValueHelper.AddDeltaToValues_PriorityForNumberGreaterThanZero(listaValoriAttuali, -deltaSpese_Totale);
                    _applicaNuovi_ValorePianificazioneConConsumi_AlMese(righeAttualiPerTupla, mese, nuoviValori);
                    //// applico i nuovi valori
                    //for (var j = 0; j < righeAttualiPerTupla.Count; j++)
                    //{
                    //    righeAttualiPerTupla[j].ValorePianificazioneConConsumi[mese] = nuoviValori[j];
                    //}

                    deltaSpese_Totale = 0;
                }
            }


            // Verifico che il deltaSpese_Totale si stato effettivamente compensanto interamente
            if (Math.Round(deltaSpese_Totale, Numbers.NumeroDecimaliImportiSpese) > 0)
            {
                // deltaSpese_Totale accumulato è ancora maggiore di zero
                // Non c'è stato abbastanza spazio nei mesi futuri per assorbirlo, lo applico come extra budget

                var extraBudgetEquidistribuitoPerRiga = Math.Round(deltaSpese_Totale / righeAttualiPerTupla.Count, Numbers.NumeroDecimaliImportiSpese);


                // codice per controlli incorciati
                //var totaleDeltaApplicato = righeAttualiPerTupla.Sum(_ => _.DeltaApplicato.Sum());                
                //if (deltaSpese_Totale != totaleDeltaApplicato)
                //{
                //    var s = "Strano!";
                //}

                for (var j = 0; j < righeAttualiPerTupla.Count; j++)
                {
                    //righeAttualiPerTupla[j].ExtraBudget = extraBudgetEquidistribuitoPerRiga;

                    var extraBudgetCalcolatoPerRiga = righeAttualiPerTupla[j].DeltaApplicato.Sum();
                    righeAttualiPerTupla[j].ExtraBudget = extraBudgetCalcolatoPerRiga;

                }
            }
        }


        /// <summary>
        /// Registra i valori calcolati per i seguenti fogli: "Pianificazione con consumi", "Spese actual", "Spese commitment", "Spese totali"
        /// </summary>
        private void Aggiorna_PianificazioneConConsumi(InfoFileReport infoFileReport, Configurazione configurazione, List<Reparto> repartiCensitiInReport, List<FornitoreCensito> fornitoriCensiti, List<RigaSpese> righeSpese, List<RigaTabellaSintesi> righeAttuali, out List<RigaLogElaborazioniTabellaSintesi> logOperazioniPianificazioneConConsumi)
        {
            var worksheetName_Source = infoFileReport.WorksheetName_Sintesi;
            var worksheetName_Destination = infoFileReport.WorksheetName_PianificazioneConConsumi;

            var ultimaRigaUsataNelFoglio = infoFileReport.EPPlusHelper.GetRowsLimit(worksheetName_Source);
            var rigaCorrente = configurazione.Sintesi_Riga_PrimaConDati;

            logOperazioniPianificazioneConConsumi = new List<RigaLogElaborazioniTabellaSintesi>();
            while (rigaCorrente <= ultimaRigaUsataNelFoglio)
            {
                #region Lettura dati identificativi della tupla
                var siglaFornitore = infoFileReport.EPPlusHelper.GetString(worksheetName_Source, rigaCorrente, configurazione.Sintesi_ColonnaSiglaFornitore);
                var reparto = infoFileReport.EPPlusHelper.GetString(worksheetName_Source, rigaCorrente, configurazione.Sintesi_ColonnaReparto);
                var objectTipologiaSpesa = infoFileReport.EPPlusHelper.GetValue(worksheetName_Source, rigaCorrente, configurazione.Sintesi_ColonnaTipologiaSpesa);

                // se tutte i valori sono nulli ho raggiunto la fine della tabella, mi fermo
                if (siglaFornitore == null && reparto == null && objectTipologiaSpesa == null)
                { break; }

                // se uno o più valori è nullo sollevo un'eccezione
                if (siglaFornitore == null || reparto == null || objectTipologiaSpesa == null)
                {
                    throw new ManagedException(
                    tipologiaErrore: TipologiaErrori.DatoMancante,
                    tipologiaCartella: TipologiaCartelle.ReportInput,
                    worksheetName: worksheetName_Source,
                    rigaCella: rigaCorrente,
                    colonnaCella: siglaFornitore == null ? configurazione.Sintesi_ColonnaSiglaFornitore
                                : reparto == null ? configurazione.Sintesi_ColonnaReparto
                                : objectTipologiaSpesa == null ? configurazione.Sintesi_ColonnaTipologiaSpesa
                                : -1,
                    nomeDatoErrore: siglaFornitore == null ? NomiDatoErrore.SiglaFornitore
                                : reparto == null ? NomiDatoErrore.NomeReparto
                                : objectTipologiaSpesa == null ? NomiDatoErrore.TipologiaSpesa
                                : NomiDatoErrore.None,
                    dato: null,
                    percorsoFile: null
                    );
                }

                // Validazione dei valori letti
                if (!ValueHelper.CastTipologieDiSpesa(objectTipologiaSpesa, out TipologieDiSpesa? tipologieDiSpesa))
                {
                    // questo errore non dovrebbe potersi verificare in quanto i valori inseribili nelle celle sono vincolati da un menù a tendina
                    throw new ManagedException(
                          tipologiaErrore: TipologiaErrori.DatoNonValido,
                          tipologiaCartella: TipologiaCartelle.ReportInput,
                          worksheetName: worksheetName_Source,
                          rigaCella: rigaCorrente,
                          colonnaCella: configurazione.Sintesi_ColonnaTipologiaSpesa,
                          nomeDatoErrore: NomiDatoErrore.TipologiaSpesa,
                          dato: objectTipologiaSpesa.ToString()
                          );
                }

                // verifico che la Sigla del fornitore sia valida
                if (!fornitoriCensiti.Any(_ => _.SiglaInReport.Equals(siglaFornitore, StringComparison.InvariantCultureIgnoreCase)))
                {
                    throw new ManagedException(
                          tipologiaErrore: TipologiaErrori.DatoNonValido,
                          tipologiaCartella: TipologiaCartelle.ReportInput,
                          worksheetName: worksheetName_Source,
                          rigaCella: rigaCorrente,
                          colonnaCella: configurazione.Sintesi_ColonnaSiglaFornitore,
                          nomeDatoErrore: NomiDatoErrore.SiglaFornitore,
                          dato: siglaFornitore
                          );
                }

                // verifico che il Reparto sia un valore valido
                if (!repartiCensitiInReport.Any(_ => _.Nome.Equals(reparto, StringComparison.InvariantCultureIgnoreCase)))
                {
                    // questo errore non dovrebbe potersi verificare in quanto i valori inseribili nelle celle sono vincolati da un menù a tendina
                    throw new ManagedException(
                          tipologiaErrore: TipologiaErrori.DatoNonValido,
                          tipologiaCartella: TipologiaCartelle.ReportInput,
                          worksheetName: worksheetName_Source,
                          rigaCella: rigaCorrente,
                          colonnaCella: configurazione.Sintesi_ColonnaReparto,
                          nomeDatoErrore: NomiDatoErrore.NomeReparto,
                          dato: reparto
                          );
                }

                var categoriaFornitori = infoFileReport.EPPlusHelper.GetString(worksheetName_Source, rigaCorrente, configurazione.Sintesi_ColonnaCategoriaFornitore);
                #endregion


                #region Selezione della riga da editare

                var rigaDaEditare = righeAttuali.FirstOrDefault(_ =>
                      _.SiglaFornitore.Equals(siglaFornitore, StringComparison.InvariantCultureIgnoreCase) &&
                      _.Reparto.Equals(reparto, StringComparison.InvariantCultureIgnoreCase) &&
                      _.TipologiaDiSpesa.Equals(tipologieDiSpesa.Value) &&
                      _.AlreadyChecked == false);
                // marco come AlreadyChecked questa riga per evitare di editarla più di una volta
                rigaDaEditare.AlreadyChecked = true;
                #endregion

                int colonnaCorrente;
                for (var mese = 0; mese <= 11; mese++)
                {
                    colonnaCorrente = configurazione.Sintesi_Colonna_PrimaRangeMesi + mese;

                    var currentValueMese = (double?)infoFileReport.EPPlusHelper.GetValue(worksheetName_Source, rigaCorrente, colonnaCorrente) ?? 0;

                    #region Aggiornamento foglio "Pianificazione con consumi"
                    // Determino valore attuale e nuovo valore per il foglio  "Pianificazione con consumi"
                    var newValueMese_per_PianificazioneConConsumi = rigaDaEditare.ValorePianificazioneConConsumi[mese];

                    // Setto il valore della cella, sostituendo gli zero con un null
                    //var valorePerLaCellaSintesi = (newValueMese_per_PianificazioneConConsumi <= 0)
                    //                        ? null
                    //                        : (double?)Math.Round(newValueMese_per_PianificazioneConConsumi, Numbers.NumeroDecimaliImportiSpese);

                    var valorePerLaCellaMese = _getValorePerCellaFoglioSpese(newValueMese_per_PianificazioneConConsumi);
                    infoFileReport.EPPlusHelper.SetValue(worksheetName_Destination, rigaCorrente, colonnaCorrente, valorePerLaCellaMese);

                    // applico l'eventuale colore di backgroud alla cella
                    if (newValueMese_per_PianificazioneConConsumi < 0)
                    {
                        var coloreBackground = BACKGROUND_COLOR_EXTRABUDGET;
                        infoFileReport.EPPlusHelper.SetBackgroundColor(worksheetName_Destination, rigaCorrente, colonnaCorrente, coloreBackground, 0);
                    }

                    // Se il valore è cambiato loggo le info sul cambiamento applicato
                    if (currentValueMese != newValueMese_per_PianificazioneConConsumi)
                    {
                        // informazioni per i log                        
                        var operazione = (newValueMese_per_PianificazioneConConsumi > currentValueMese)
                                        ? "Aggiunte"    // è stato aggiunto qualcosa al valore esistente
                                        : "Sottratte";  // è stato sottratto qualcosa al valore esistente
                        var valore = Math.Abs(currentValueMese - newValueMese_per_PianificazioneConConsumi);
                        var simboloValore = (rigaDaEditare.TipologiaDiSpesa == TipologieDiSpesa.AdOre)
                                        ? "ore"
                                        : "euro";
                        logOperazioniPianificazioneConConsumi.Add(
                                new RigaLogElaborazioniTabellaSintesi(
                                        reparto: reparto,
                                        operazione: operazione,
                                        valore: valore,
                                        simboloValore: simboloValore,
                                        mese: mese + 1,
                                        fornitore: siglaFornitore,
                                        riga: rigaCorrente
                                        ));
                    }
                    #endregion
                }

                if (rigaDaEditare.ExtraBudget > 0)
                {
                    // Setto il valore di extra budget nella colonna predisposta
                    colonnaCorrente = configurazione.Sintesi_Colonna_PrimaRangeMesi + 12;
                    infoFileReport.EPPlusHelper.SetValue(worksheetName_Destination, rigaCorrente, colonnaCorrente, rigaDaEditare.ExtraBudget);

                    // applico il colore di backgroud alla cella
                    infoFileReport.EPPlusHelper.SetBackgroundColor(worksheetName_Destination, rigaCorrente, colonnaCorrente, BACKGROUND_COLOR_EXTRABUDGET, 0);
                }

                // vado alla riga successiva
                rigaCorrente++;
            }
        }



        ///// <summary>
        ///// Registra i valori calcolati per i seguenti fogli: "Pianificazione con consumi", "Spese actual", "Spese commitment", "Spese totali"
        ///// </summary>
        //private void Aggiorna_FogliGemelli(InfoFileReport infoFileReport, Configurazione configurazione, List<Reparto> repartiCensitiInReport, List<FornitoreCensito> fornitoriCensiti, List<TuplaSpeseTabellaSintesi> setTupleSpese, List<RigaTabellaSintesi> righeAttuali, out List<RigaLogElaborazioniTabellaSintesi> logOperazioniPianificazioneConConsumi)
        //{
        //    var worksheetName_Source = infoFileReport.WorksheetName_Sintesi;

        //    var ultimaRigaUsataNelFoglio = infoFileReport.EPPlusHelper.GetRowsLimit(worksheetName_Source);
        //    var rigaCorrente = configurazione.Sintesi_Riga_PrimaConDati;

        //    logOperazioniPianificazioneConConsumi = new List<RigaLogElaborazioniTabellaSintesi>();
        //    while (rigaCorrente <= ultimaRigaUsataNelFoglio)
        //    {
        //        #region Lettura dati identificativi della tupla
        //        var siglaFornitore = infoFileReport.EPPlusHelper.GetString(worksheetName_Source, rigaCorrente, configurazione.Sintesi_ColonnaSiglaFornitore);
        //        var reparto = infoFileReport.EPPlusHelper.GetString(worksheetName_Source, rigaCorrente, configurazione.Sintesi_ColonnaReparto);
        //        var objectTipologiaSpesa = infoFileReport.EPPlusHelper.GetValue(worksheetName_Source, rigaCorrente, configurazione.Sintesi_ColonnaTipologiaSpesa);

        //        // se tutte i valori sono nulli ho raggiunto la fine della tabella, mi fermo
        //        if (siglaFornitore == null && reparto == null && objectTipologiaSpesa == null)
        //        { break; }

        //        // se uno o più valori è nullo sollevo un'eccezione
        //        if (siglaFornitore == null || reparto == null || objectTipologiaSpesa == null)
        //        {
        //            throw new ManagedException(
        //            tipologiaErrore: TipologiaErrori.DatoMancante,
        //            tipologiaCartella: TipologiaCartelle.ReportInput,
        //            worksheetName: worksheetName_Source,
        //            rigaCella: rigaCorrente,
        //            colonnaCella: siglaFornitore == null ? configurazione.Sintesi_ColonnaSiglaFornitore
        //                        : reparto == null ? configurazione.Sintesi_ColonnaReparto
        //                        : objectTipologiaSpesa == null ? configurazione.Sintesi_ColonnaTipologiaSpesa
        //                        : -1,
        //            nomeDatoErrore: siglaFornitore == null ? NomiDatoErrore.SiglaFornitore
        //                        : reparto == null ? NomiDatoErrore.NomeReparto
        //                        : objectTipologiaSpesa == null ? NomiDatoErrore.TipologiaSpesa
        //                        : NomiDatoErrore.None,
        //            dato: null,
        //            percorsoFile: null
        //            );
        //        }

        //        // Validazione dei valori letti
        //        if (!ValueHelper.CastTipologieDiSpesa(objectTipologiaSpesa, out TipologieDiSpesa? tipologieDiSpesa))
        //        {
        //            // questo errore non dovrebbe potersi verificare in quanto i valori inseribili nelle celle sono vincolati da un menù a tendina
        //            throw new ManagedException(
        //                  tipologiaErrore: TipologiaErrori.DatoNonValido,
        //                  tipologiaCartella: TipologiaCartelle.ReportInput,
        //                  worksheetName: worksheetName_Source,
        //                  rigaCella: rigaCorrente,
        //                  colonnaCella: configurazione.Sintesi_ColonnaTipologiaSpesa,
        //                  nomeDatoErrore: NomiDatoErrore.TipologiaSpesa,
        //                  dato: objectTipologiaSpesa.ToString()
        //                  );
        //        }

        //        // verifico che la Sigla del fornitore sia valida
        //        if (!fornitoriCensiti.Any(_ => _.SiglaInReport.Equals(siglaFornitore, StringComparison.InvariantCultureIgnoreCase)))
        //        {
        //            throw new ManagedException(
        //                  tipologiaErrore: TipologiaErrori.DatoNonValido,
        //                  tipologiaCartella: TipologiaCartelle.ReportInput,
        //                  worksheetName: worksheetName_Source,
        //                  rigaCella: rigaCorrente,
        //                  colonnaCella: configurazione.Sintesi_ColonnaSiglaFornitore,
        //                  nomeDatoErrore: NomiDatoErrore.SiglaFornitore,
        //                  dato: siglaFornitore
        //                  );
        //        }

        //        // verifico che il Reparto sia un valore valido
        //        if (!repartiCensitiInReport.Any(_ => _.Nome.Equals(reparto, StringComparison.InvariantCultureIgnoreCase)))
        //        {
        //            // questo errore non dovrebbe potersi verificare in quanto i valori inseribili nelle celle sono vincolati da un menù a tendina
        //            throw new ManagedException(
        //                  tipologiaErrore: TipologiaErrori.DatoNonValido,
        //                  tipologiaCartella: TipologiaCartelle.ReportInput,
        //                  worksheetName: worksheetName_Source,
        //                  rigaCella: rigaCorrente,
        //                  colonnaCella: configurazione.Sintesi_ColonnaReparto,
        //                  nomeDatoErrore: NomiDatoErrore.NomeReparto,
        //                  dato: reparto
        //                  );
        //        }

        //        var categoriaFornitori = infoFileReport.EPPlusHelper.GetString(worksheetName_Source, rigaCorrente, configurazione.Sintesi_ColonnaCategoriaFornitore);
        //        #endregion

        //        #region Selezione della riga da editare
        //        var rigaDaEditare = righeAttuali.FirstOrDefault(_ =>
        //              _.SiglaFornitore.Equals(siglaFornitore, StringComparison.InvariantCultureIgnoreCase) &&
        //              _.Reparto.Equals(reparto, StringComparison.InvariantCultureIgnoreCase) &&
        //              _.TipologiaDiSpesa.Equals(tipologieDiSpesa.Value) &&
        //              _.AlreadyChecked == false);
        //        // marco come AlreadyChecked questa riga per evitare di editarla più di una volta
        //        rigaDaEditare.AlreadyChecked = true;
        //        #endregion

        //        for (var mese = 0; mese <= 11; mese++)
        //        {
        //            var colonnaCorrente = configurazione.Sintesi_Colonna_PrimaRangeMesi + mese;

        //            var currentValueMese = (double?)infoFileReport.EPPlusHelper.GetValue(worksheetName_Source, rigaCorrente, colonnaCorrente) ?? 0;


        //            #region Aggiornamento foglio "Pianificazione con consumi"
        //            // Determino valore attuale e nuovo valore per il foglio  "Pianificazione con consumi"
        //            var newValueMese_per_PianificazioneConConsumi = rigaDaEditare.NewValues[mese];

        //            // Setto il valore della cella, sostituendo gli zero con un null
        //            var valorePerLaCellaSintesi = (newValueMese_per_PianificazioneConConsumi <= 0)
        //                                    ? null
        //                                    : (double?)Math.Round(newValueMese_per_PianificazioneConConsumi, 2);
        //            infoFileReport.EPPlusHelper.SetValue(infoFileReport.WorksheetName_PianificazioneConConsumi, rigaCorrente, colonnaCorrente, valorePerLaCellaSintesi);

        //            // applico l'eventuale colore di backgroud alla cella
        //            if (newValueMese_per_PianificazioneConConsumi < 0)
        //            {
        //                var coloreBackground = BACKGROUND_COLOR_VALUE_OUT_OF_PLANNING;
        //                infoFileReport.EPPlusHelper.SetBackgroundColor(infoFileReport.WorksheetName_PianificazioneConConsumi, rigaCorrente, colonnaCorrente, coloreBackground, 0);
        //            }

        //            // Se il valore è cambiato loggo le info sul cambiamento applicato
        //            if (currentValueMese != newValueMese_per_PianificazioneConConsumi)
        //            {
        //                // informazioni per i log                        
        //                var operazione = (currentValueMese - newValueMese_per_PianificazioneConConsumi < 0)
        //                            ? "Over budget"         // il valore è andato in negativo
        //                            : (newValueMese_per_PianificazioneConConsumi > currentValueMese)
        //                                    ? "Aggiunte"    // è stato aggiunto qualcosa al valore esistente
        //                                    : "Sottratte";  // è stato sottratto qualcosa al valore esistente
        //                var valore = Math.Abs(currentValueMese - newValueMese_per_PianificazioneConConsumi);

        //                var simboloValore = (rigaDaEditare.TipologiaDiSpesa == TipologieDiSpesa.AdOre) ? "ore" : "euro";

        //                logOperazioniPianificazioneConConsumi.Add(
        //                        new RigaLogElaborazioniTabellaSintesi(
        //                                reparto: reparto,
        //                                operazione: operazione,
        //                                valore: valore,
        //                                simboloValore: simboloValore,
        //                                mese: mese + 1,
        //                                fornitore: siglaFornitore,
        //                                riga: rigaCorrente
        //                                ));
        //            }
        //            #endregion

        //            #region Aggiornamento fogli gemelli
        //            var tuplaSource = setTupleSpese.SingleOrDefault(_ =>
        //                      _.SiglaFornitore.Equals(siglaFornitore, StringComparison.InvariantCultureIgnoreCase) &&
        //                      _.Reparto.Equals(reparto, StringComparison.InvariantCultureIgnoreCase) &&
        //                      _.TipologiaDiSpesa.Equals(tipologieDiSpesa.Value));

        //            // non tutte le righe presenti in "Sintesi" potrebbero avere una spesa, quindi non essere presenti nella lista tuplaSource
        //            if (tuplaSource != null)
        //            {
        //                // Aggiorno foglio "Spese Actual"
        //                var valorePerLaCella_SpeseActual = _getValorePerCellaFoglioSpese(tuplaSource.SpeseActualPerMese[mese]);
        //                infoFileReport.EPPlusHelper.SetValue(infoFileReport.WorksheetName_SpeseActual, rigaCorrente, colonnaCorrente, valorePerLaCella_SpeseActual);

        //                // Aggiorno foglio "Spese Comittment"
        //                var valorePerLaCella_SpeseComittment = _getValorePerCellaFoglioSpese(tuplaSource.SpeseCommittedPerMese[mese]);
        //                infoFileReport.EPPlusHelper.SetValue(infoFileReport.WorksheetName_SpeseComittment, rigaCorrente, colonnaCorrente, valorePerLaCella_SpeseComittment);

        //                // Aggiorno foglio "Spese Totali"
        //                var valorePerLaCella_SpeseTotali = _getValorePerCellaFoglioSpese(tuplaSource.SpeseActualPerMese[mese] + tuplaSource.SpeseCommittedPerMese[mese]);
        //                infoFileReport.EPPlusHelper.SetValue(infoFileReport.WorksheetName_SpeseTotali, rigaCorrente, colonnaCorrente, valorePerLaCella_SpeseTotali);
        //            }
        //            #endregion
        //        }

        //        // vado alla riga successiva
        //        rigaCorrente++;
        //    }
        //}




        /// <summary>
        /// 
        /// </summary>
        private void RegistraOperazioniSuPianificazioneConConsumi_SuFoglioLogModifiche(InfoFileReport infoFileReport, List<RigaLogElaborazioniTabellaSintesi> cambiamentiAllaTabellaSintesi, int periodo)
        {
            var cambiamentiPassato = cambiamentiAllaTabellaSintesi.Where(_ => _.Mese < periodo).ToList();
            _registraCambiamentiSuFoglioLogModificheTabella(infoFileReport, cambiamentiPassato, 1);

            var cambiamentiFuturo = cambiamentiAllaTabellaSintesi.Where(_ => _.Mese >= periodo).ToList();
            _registraCambiamentiSuFoglioLogModificheTabella(infoFileReport, cambiamentiFuturo, 8);
        }
        private void _registraCambiamentiSuFoglioLogModificheTabella(InfoFileReport infoFileReport, List<RigaLogElaborazioniTabellaSintesi> cambiamentiAllaTabellaSintesi, int primaColonnaTabella)
        {
            var worksheetName = infoFileReport.WorksheetName_SintesiLogModifiche;  // Sintesi-Log modifiche
            var rigaCorrente = 3; // si comincia dalla riga numero 3

            // registro i dati correnti
            foreach (var cambiamento in cambiamentiAllaTabellaSintesi.OrderBy(_ => _.Reparto)
                            .ThenBy(_ => _.Mese)
                            .ThenBy(_ => _.Fornitore)
                            .ThenBy(_ => _.SimboloValore)
                            .ThenBy(_ => _.Operazione))
            {
                var colonnaCorrente = primaColonnaTabella;

                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaCorrente, colonnaCorrente++, cambiamento.Reparto);
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaCorrente, colonnaCorrente++, cambiamento.Operazione);
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaCorrente, colonnaCorrente++, Math.Round(cambiamento.Valore, Numbers.NumeroDecimaliImportiSpese) + " " + cambiamento.SimboloValore);
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaCorrente, colonnaCorrente++, new DateTime(2024, cambiamento.Mese, 1).ToString("MMMM"));
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaCorrente, colonnaCorrente++, cambiamento.Fornitore);
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaCorrente, colonnaCorrente++, cambiamento.Riga);

                rigaCorrente++;
            }

            // Azzeramento delle righe mancanti
            var ultimaRigaUsataNelFoglio = infoFileReport.EPPlusHelper.GetRowsLimit(worksheetName);
            while (rigaCorrente <= ultimaRigaUsataNelFoglio)
            {
                var colonnaCorrente = primaColonnaTabella;

                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaCorrente, colonnaCorrente++, "");
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaCorrente, colonnaCorrente++, "");
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaCorrente, colonnaCorrente++, "");
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaCorrente, colonnaCorrente++, "");
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaCorrente, colonnaCorrente++, "");
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaCorrente, colonnaCorrente++, "");

                rigaCorrente++;
            }
        }

        private void Aggiorna_FogliSpese(InfoFileReport infoFileReport, Configurazione configurazione, List<TuplaSpesePerAnno> tupleSpese)
        {
            _svuotaTabella_FogliSpese(infoFileReport, configurazione);
            _scriviTupleSplesa(infoFileReport, configurazione, tupleSpese);
        }
        private void _svuotaTabella_FogliSpese(InfoFileReport infoFileReport, Configurazione configurazione)
        {
            // Lista dei fogli da svuotare
            var worksheetNameTo_List = new List<string>
                {
                infoFileReport.WorksheetName_SpeseActual,
                infoFileReport.WorksheetName_SpeseComittment,
                infoFileReport.WorksheetName_SpeseTotali
                };

            foreach (var worksheetNameTo in worksheetNameTo_List)
            {
                // svuoto il contenuto del foglio
                var ultimaRigaUsataNelFoglio = infoFileReport.EPPlusHelper.GetRowsLimit(worksheetNameTo);
                infoFileReport.EPPlusHelper.CleanCellsContent(worksheetNameTo, configurazione.Sintesi_Riga_PrimaConDati, 1, ultimaRigaUsataNelFoglio, configurazione.Sintesi_ColonnaOverBudget);
            }
        }
        private void _scriviTupleSplesa(InfoFileReport infoFileReport, Configurazione configurazione, List<TuplaSpesePerAnno> tupleSpese)
        {
            // Lista dei fogli da svuotare
            var worksheetNameTo_List = new List<string>
                {
                infoFileReport.WorksheetName_SpeseActual,
                infoFileReport.WorksheetName_SpeseComittment,
                infoFileReport.WorksheetName_SpeseTotali
                };

            var rigaCorrente = configurazione.Spese_Riga_PrimaConDati;
            foreach (var tuplaSpese in tupleSpese.OrderBy(_ => _.SiglaFornitore).ThenBy(_ => _.Reparto).ThenBy(_ => _.TipologiaDiSpesa))
            {
                foreach (var worksheetNameTo in worksheetNameTo_List)
                {
                    infoFileReport.EPPlusHelper.SetValue(worksheetNameTo, rigaCorrente, 1, tuplaSpese.SiglaFornitore);
                    infoFileReport.EPPlusHelper.SetValue(worksheetNameTo, rigaCorrente, 2, tuplaSpese.Reparto);
                    infoFileReport.EPPlusHelper.SetValue(worksheetNameTo, rigaCorrente, 3, tuplaSpese.TipologiaDiSpesa);
                    for (var mese = 0; mese <= 11; mese++)
                    {
                        var colonnaCorrente = configurazione.Spese_Colonna_PrimaRangeMesi + mese;

                        double? value = 0;
                        if (worksheetNameTo.Equals(infoFileReport.WorksheetName_SpeseActual))
                            value = _getValorePerCellaFoglioSpese(tuplaSpese.SpeseActualPerMese[mese]);
                        else if (worksheetNameTo.Equals(infoFileReport.WorksheetName_SpeseComittment))
                            value = _getValorePerCellaFoglioSpese(tuplaSpese.SpeseCommittedPerMese[mese]);
                        else if (worksheetNameTo.Equals(infoFileReport.WorksheetName_SpeseTotali))
                            value = _getValorePerCellaFoglioSpese(tuplaSpese.SpeseActualPerMese[mese] + tuplaSpese.SpeseCommittedPerMese[mese]);

                        infoFileReport.EPPlusHelper.SetValue(worksheetNameTo, rigaCorrente, colonnaCorrente, value);
                    }
                }
                rigaCorrente++;// next tupla
            }

        }


        /// <summary>
        /// Setta il colore delle celle dei mesi passati con un colore di background
        /// </summary>
        private void SettaColoreBackgroundAlleCellePeriodoPassato(InfoFileReport infoFileReport, Configurazione configurazione, int periodo, int numeroRighe)
        {
            if (numeroRighe == 0)
            { return; }

            // Lista dei fogli in cui settare il colore di backgroud
            var worksheetNameTo_List = new List<string>
                {
                infoFileReport.WorksheetName_Sintesi,
                infoFileReport.WorksheetName_PianificazioneConConsumi,
                infoFileReport.WorksheetName_SpeseActual,
                infoFileReport.WorksheetName_SpeseComittment,
                infoFileReport.WorksheetName_SpeseTotali
                };

            foreach (var worksheetName in worksheetNameTo_List)
            {
                var primaRigaConDati = (worksheetName.Equals(infoFileReport.WorksheetName_Sintesi) || worksheetName.Equals(infoFileReport.WorksheetName_PianificazioneConConsumi))
                        ? configurazione.Sintesi_Riga_PrimaConDati
                        : configurazione.Spese_Riga_PrimaConDati;
                var primaColonnaMesi = (worksheetName.Equals(infoFileReport.WorksheetName_Sintesi) || worksheetName.Equals(infoFileReport.WorksheetName_PianificazioneConConsumi))
                        ? configurazione.Sintesi_Colonna_PrimaRangeMesi
                        : configurazione.Spese_Colonna_PrimaRangeMesi;

                // individuo l'ampiezza delle righe
                var rowFrom = primaRigaConDati;
                var rowTo = infoFileReport.EPPlusHelper.GetLastUsedRowForColumn(worksheetName, primaRigaConDati, 1);
                if (rowTo == 0)
                { continue; /*tabella vuota*/}

                // setto il backgroud di tutte le celle nei mesi "Passati"
                var colFrom = primaColonnaMesi;
                var colTo = (periodo != 1)
                           ? colFrom + (periodo - 1) - 1
                           : colFrom + 11; // in caso di periodo=1, siamo a gennaio e l'intero periodo è da considerarsi passato
                infoFileReport.EPPlusHelper.SetBackgroundColor(worksheetName, rowFrom, colFrom, rowTo, colTo, BACKGROUND_COLOR_PAST_MONTHS, 0);

                // setto il backgroud di tutte le celle nei mesi "Futuri"
                if (periodo != 1) // in caso di periodo = 1 non ci sono mesi futuri
                {
                    colFrom = primaColonnaMesi + (periodo - 1);
                    colTo = primaColonnaMesi + 12 - 1;
                    infoFileReport.EPPlusHelper.SetBackgroundColor(worksheetName, rowFrom, colFrom, rowTo, colTo, BACKGROUND_COLOR_FUTURE_MONTHS, 0);
                }
            }
        }



        private void _applicaNuovi_ValorePianificazioneConConsumi_AlMese(List<RigaTabellaSintesi> righeAttualiPerTupla, int mese, List<double> nuoviValori)
        {
            // applico i nuovi valori
            for (var j = 0; j < righeAttualiPerTupla.Count; j++)
            {
                // nuovo valore per la tabella "Pianificazione con consumi"
                righeAttualiPerTupla[j].ValorePianificazioneConConsumi[mese] = nuoviValori[j];

                // registro la variazione di valore
                var deltaApplicato = nuoviValori[j] - righeAttualiPerTupla[j].BudgetStanziato[mese];
                righeAttualiPerTupla[j].DeltaApplicato[mese] += deltaApplicato;
            }
        }

        private double? _getValorePerCellaFoglioSpese(double value)
        {
            var roundedValue = Math.Round(value, Numbers.NumeroDecimaliImportiSpese);
            return (roundedValue <= 0)
                    ? null
                    : (double?)roundedValue;
        }
    }
}