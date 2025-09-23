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
    /// Funzionalità in comune tra più steps per l'nserimento dei nuovi fornitori ricevuti come input
    /// </summary>
    internal abstract class Step_Inserimento_NuoviFornitori_Base : Step_Base
    {
        internal void InserimentoNuoviFornitori(InfoFileReport infoFileReport, Configurazione configurazione, List<FornitoreCensito> fornitoriCensiti, List<Reparto> repartiCensitiInReport, List<string> categorieFornitori, List<FornitoreCensito> fornitoriDaAggiungere)
        {
            VerificaCorrettezzaDeiNuoviFornitori(infoFileReport, fornitoriCensiti, repartiCensitiInReport, categorieFornitori, fornitoriDaAggiungere);
            InserimentoNuoviFornitori_AnagraficaFornitori(infoFileReport, configurazione, fornitoriDaAggiungere);
            InserimentoNuoviFornitori_ListaDati(infoFileReport, configurazione, repartiCensitiInReport, fornitoriDaAggiungere);
            InserimentoNuoviFornitori_BudgetStudiIpotesi1(infoFileReport, configurazione, fornitoriCensiti, fornitoriDaAggiungere);
            InserimentoNuoviFornitori_Reportistica(infoFileReport, configurazione, fornitoriCensiti, fornitoriDaAggiungere);
            InserimentoNuoviFornitori_ReportisticaPerCategoria(infoFileReport, configurazione, fornitoriDaAggiungere);

            // merge dei nuovi fornitori inseriri tra quelli censiti
            fornitoriCensiti.AddRange(fornitoriDaAggiungere.Where(_ => !_.PresenteSoloInListaDati));
        }

        private void VerificaCorrettezzaDeiNuoviFornitori(InfoFileReport infoFileReport, List<FornitoreCensito> fornitoriCensiti, List<Reparto> repartiCensitiInReport, List<string> categorieFornitori, List<FornitoreCensito> fornitoriDaAggiungere)
        {
            // N.B.dai fornitori da aggiugnere vengono esclusi dai controlli i fornitori con PresenteSoloInListaDati a true
            foreach (var fornitoreDaAggiungere in fornitoriDaAggiungere.Where(_ => !_.PresenteSoloInListaDati))
            {
                // caso di fornitore con Sigla non univoca
                if (fornitoriCensiti.Any(_ => _.SiglaInReport.Equals(fornitoreDaAggiungere.SiglaInReport, StringComparison.InvariantCultureIgnoreCase)))
                {
                    throw new ManagedException(
                         tipologiaErrore: TipologiaErrori.DatoNonUnivoco,
                         tipologiaCartella: TipologiaCartelle.ReportInput,
                         nomeDatoErrore: NomiDatoErrore.SiglaFornitore,
                         worksheetName: infoFileReport.WorksheetName_AnagraficaFornitori,
                         rigaCella: null,
                         colonnaCella: null,
                         dato: fornitoreDaAggiungere.SiglaInReport,
                         percorsoFile: null
                         );
                }
                // caso di fornitore con Nome non univoco
                if (fornitoriCensiti.Any(_ => _.HasThisName(fornitoreDaAggiungere.NomeSuController)))
                {
                    throw new ManagedException(
                         tipologiaErrore: TipologiaErrori.DatoNonUnivoco,
                         tipologiaCartella: TipologiaCartelle.ReportInput,
                         nomeDatoErrore: NomiDatoErrore.NomeFornitore,
                         worksheetName: infoFileReport.WorksheetName_AnagraficaFornitori,
                         rigaCella: null,
                         colonnaCella: null,
                         dato: fornitoreDaAggiungere.NomeSuController,
                         percorsoFile: null
                         );
                }

                // caso (limite) ma possibile, si verifica che tutti i reparti a cui è stato dato un costo custom
                // tramite l'interfaccia siano effettivamente reparti validi per il file Report
                var repartiConCostiCustom = fornitoreDaAggiungere.GetRepartiConCostiCustom();
                if (repartiConCostiCustom != null)
                {
                    foreach (var repartoConCostiCustom in repartiConCostiCustom)
                    {
                        if (!repartiCensitiInReport.Any(_ => _.Nome.Equals(repartoConCostiCustom, StringComparison.InvariantCultureIgnoreCase)))
                        {
                            // il nome del reparto con costo custon non è censito in Report
                            throw new ManagedException(
                                 tipologiaErrore: TipologiaErrori.DatoNonValido,
                                 tipologiaCartella: TipologiaCartelle.ReportInput,
                                 nomeDatoErrore: NomiDatoErrore.NomeReparto,
                                 worksheetName: infoFileReport.WorksheetName_ListaDati,
                                 rigaCella: null,
                                 colonnaCella: null,
                                 dato: repartoConCostiCustom,
                                 percorsoFile: null
                                 );
                        }
                    }
                }

                // Verifico che la categoria assegnata al fornitore sia tra quelle censite in Report
                if (!categorieFornitori.Any(_ => _.Equals(fornitoreDaAggiungere.Categoria, StringComparison.InvariantCultureIgnoreCase)))
                {
                    throw new ManagedException(
                         tipologiaErrore: TipologiaErrori.DatoNonValido,
                         tipologiaCartella: TipologiaCartelle.ReportInput,
                         nomeDatoErrore: NomiDatoErrore.CategoriaFornitore,
                         worksheetName: infoFileReport.WorksheetName_ListaDati,
                         rigaCella: null,
                         colonnaCella: null,
                         dato: fornitoreDaAggiungere.Categoria,
                         percorsoFile: null
                         );
                }
            }
        }
        private void InserimentoNuoviFornitori_AnagraficaFornitori(InfoFileReport infoFileReport, Configurazione configurazione, List<FornitoreCensito> fornitoriDaAggiungere)
        {
            var worksheetName = infoFileReport.WorksheetName_AnagraficaFornitori; // Anagrafica fornitori

            // aggiungo i nuovi fornitori in coda alla tabella
            foreach (var fornitoreDaAggiungere in fornitoriDaAggiungere)
            {
                var rigaInCuiAggiungereNuovoFornitore = configurazione.AnagraficaFornitori_PrimaRigaFornitori;
                infoFileReport.EPPlusHelper.AggiungiRighe(worksheetName, rigaInCuiAggiungereNuovoFornitore);
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaInCuiAggiungereNuovoFornitore, configurazione.AnagraficaFornitori_ColonnaSigle, fornitoreDaAggiungere.SiglaInReport, true);
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaInCuiAggiungereNuovoFornitore, configurazione.AnagraficaFornitori_ColonnaNomi, fornitoreDaAggiungere.NomeSuController, true);
            }
        }
        private void InserimentoNuoviFornitori_ListaDati(InfoFileReport infoFileReport, Configurazione configurazione, List<Reparto> repartiCensitiInReport, List<FornitoreCensito> fornitoriDaAggiungere)
        {
            var worksheetName = infoFileReport.WorksheetName_ListaDati; // Lista dati

            #region lettura dei reparti dalla tabella dei costi specifici per reparto
            var rigaReparti = configurazione.ListaDati_RigaReparti;
            var nomiRepartiSuReport = new List<string>();
            var colonnaCorrente = configurazione.ListaDati_PrimaColonnaReparti;
            while (true)
            {
                var nomeReparto = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaReparti, colonnaCorrente);
                if (string.IsNullOrWhiteSpace(nomeReparto)) // mi interrompo con l'ultima cella con eventuale valore blank
                { break; }

                // Mi attengo la nome del reparto già censito
                var repartoCensitoInReport = repartiCensitiInReport.SingleOrDefault(_ => _.Nome.Equals(nomeReparto, StringComparison.InvariantCultureIgnoreCase));
                nomiRepartiSuReport.Add(repartoCensitoInReport.Nome);
                colonnaCorrente++;
            }
            var ultimaColonnaReparti = colonnaCorrente - 1;
            #endregion

            #region aggiorno le righe delle due tabelle
            foreach (var fornitoreDaAggiungere in fornitoriDaAggiungere)
            {
                // Francesco il 03/11/2023 ci ha chiesto che fosse possibile avere in "Lista dati" la sigla di un fornitore non ancora presente in "Anagrafica fornitori"

                // Cerco la sigla del fornitore nella tabella, potrebbe essere infatti già presente.
                var rigaInCuiAggiungereNuovoFornitore = infoFileReport.EPPlusHelper.GetFirstRowWithSpecificValue(worksheetName, configurazione.ListaDati_PrimaRigaFornitori, configurazione.ListaDati_ColonnaSiglaFornitoriTabellaSX, fornitoreDaAggiungere.SiglaInReport);
                if (rigaInCuiAggiungereNuovoFornitore < 1)
                {
                    // non ho trovato il fornitore, non è già presente in "Lista dati", lo inserisco alla fine
                    rigaInCuiAggiungereNuovoFornitore = infoFileReport.EPPlusHelper.GetFirstEmptyRow(
                        worksheetName: worksheetName,
                        rowToStartSearchFrom: configurazione.ListaDati_PrimaRigaFornitori,
                        colToBeChecked: configurazione.ListaDati_ColonnaSiglaFornitoriTabellaSX
                        );
                }

                // setto i valori della tabella di sinistra
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaInCuiAggiungereNuovoFornitore, configurazione.ListaDati_ColonnaSiglaFornitoriTabellaSX, fornitoreDaAggiungere.SiglaInReport, true);
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaInCuiAggiungereNuovoFornitore, configurazione.ListaDati_ColonnaCostiOrariStandard, fornitoreDaAggiungere.CostoOrarioStandard, true);
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaInCuiAggiungereNuovoFornitore, configurazione.ListaDati_ColonnaCategoriaFornitori, fornitoreDaAggiungere.Categoria, true);

                // setto i valori della tabella di destra                
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaInCuiAggiungereNuovoFornitore, configurazione.ListaDati_ColonnaSiglaFornitoriTabellaDX, fornitoreDaAggiungere.SiglaInReport, true);
                var offSetColannaCosto = 0;
                foreach (var nomeReparto in nomiRepartiSuReport)
                {
                    if (fornitoreDaAggiungere.HasRepartCostoOrarioCustom(nomeReparto))
                    {
                        // setto i valori dei prezzi custom nella tabella di destra
                        var costoSpeficifoPerReparto = fornitoreDaAggiungere.GetCostoOrarioPerReparto(nomeReparto);
                        infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaInCuiAggiungereNuovoFornitore, configurazione.ListaDati_PrimaColonnaReparti + offSetColannaCosto, costoSpeficifoPerReparto, true);
                    }
                    offSetColannaCosto++;
                }
            }
            #endregion
        }
        private void InserimentoNuoviFornitori_BudgetStudiIpotesi1(InfoFileReport infoFileReport, Configurazione configurazione, List<FornitoreCensito> fornitoriCensiti, List<FornitoreCensito> fornitoriDaAggiungere)
        {
            var worksheetName = infoFileReport.WorksheetName_BudgetStudiIpotesi; // BUDGET STUDI -  IPOTESI 1

            var _listaFornitori = new List<FornitoreCensito>();

            var ultimaRigaUsataNelFoglio = infoFileReport.EPPlusHelper.GetRowsLimit(worksheetName);
            var rigaCorrente = configurazione.IpotesiStudio_PrimaRigaFornitori;
            // mi posiziono sulla prima riga disponibile
            while (rigaCorrente <= ultimaRigaUsataNelFoglio)
            {
                var siglaFornitore = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.IpotesiStudio_ColonnaSigle);

                // se tutte i valori sono nulli ho raggiunto la fine della tabella, mi fermo
                if (siglaFornitore == null)
                { break; }

                if (!fornitoriCensiti.Any(_ => _.SiglaInReport.Equals(siglaFornitore, StringComparison.InvariantCultureIgnoreCase)))
                {
                    //Sigla fornitore presente in "BUDGET STUDI -  IPOTESI 1" ma non censito
                    throw new ManagedException(
                        tipologiaErrore: TipologiaErrori.DatoNonValido,
                        tipologiaCartella: TipologiaCartelle.ReportInput,
                        nomeDatoErrore: NomiDatoErrore.SiglaFornitore,
                        worksheetName: worksheetName,
                        rigaCella: rigaCorrente,
                        colonnaCella: configurazione.IpotesiStudio_ColonnaSigle,
                        dato: siglaFornitore,
                        percorsoFile: null
                        );
                }

                // vado alla riga successiva
                rigaCorrente++;
            }



            #region aggiorno le righe delle due tabelle
            foreach (var fornitoreDaAggiungere in fornitoriDaAggiungere)
            {
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaCorrente, configurazione.IpotesiStudio_ColonnaSigle, fornitoreDaAggiungere.SiglaInReport, true);
                rigaCorrente++;
            }
            #endregion
        }
        private void InserimentoNuoviFornitori_Reportistica(InfoFileReport infoFileReport, Configurazione configurazione, List<FornitoreCensito> fornitoriCensiti, List<FornitoreCensito> fornitoriDaAggiungere)
        {
            var worksheetName = infoFileReport.WorksheetName_Reportistica; // REPORTISTICA
            var _listaFornitori = new List<FornitoreCensito>();

            var ultimaRigaUsataNelFoglio = infoFileReport.EPPlusHelper.GetRowsLimit(worksheetName);
            var rigaCorrente = configurazione.Reportistica_PrimaRigaFornitori;

            // mi posiziono sulla prima riga disponibile
            while (rigaCorrente <= ultimaRigaUsataNelFoglio)
            {
                var siglaFornitoreSx = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.Reportistica_ColonnaSigle_SX);
                var siglaFornitoreDx = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.Reportistica_ColonnaSigle_DX);

                // se tutte i valori sono nulli ho raggiunto la fine della tabella, mi fermo
                if (siglaFornitoreSx == null && siglaFornitoreDx == null)
                { break; }

                if (fornitoriCensiti.SingleOrDefault(_ => _.SiglaInReport.Equals(siglaFornitoreSx, StringComparison.InvariantCultureIgnoreCase)) == null)
                {
                    //Sigla fornitore presente in "REPORTISTICA" colonna di Sinistra ma non presente in "Lista dati"
                    throw new ManagedException(
                        tipologiaErrore: TipologiaErrori.DatoNonValido,
                        tipologiaCartella: TipologiaCartelle.ReportInput,
                        nomeDatoErrore: NomiDatoErrore.SiglaFornitore,
                        worksheetName: worksheetName,
                        rigaCella: rigaCorrente,
                        colonnaCella: configurazione.Reportistica_ColonnaSigle_SX,
                        dato: siglaFornitoreSx,
                        percorsoFile: null
                        );
                }

                if (fornitoriCensiti.SingleOrDefault(_ => _.SiglaInReport.Equals(siglaFornitoreDx, StringComparison.InvariantCultureIgnoreCase)) == null)
                {
                    //Sigla fornitore presente in "REPORTISTICA" colonna di Destra ma non presente in "Lista dati"
                    throw new ManagedException(
                        tipologiaErrore: TipologiaErrori.DatoNonValido,
                        tipologiaCartella: TipologiaCartelle.ReportInput,
                        nomeDatoErrore: NomiDatoErrore.SiglaFornitore,
                        worksheetName: worksheetName,
                        rigaCella: rigaCorrente,
                        colonnaCella: configurazione.Reportistica_ColonnaSigle_DX,
                        dato: siglaFornitoreDx,
                        percorsoFile: null
                        );
                }

                // vado alla riga successiva
                rigaCorrente++;
            }

            #region aggiorno le righe delle due tabelle
            foreach (var fornitoreDaAggiungere in fornitoriDaAggiungere)
            {
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaCorrente, configurazione.Reportistica_ColonnaSigle_SX, fornitoreDaAggiungere.SiglaInReport, true);
                infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaCorrente, configurazione.Reportistica_ColonnaSigle_DX, fornitoreDaAggiungere.SiglaInReport, true);
                rigaCorrente++;
            }
            #endregion
        }
        private void InserimentoNuoviFornitori_ReportisticaPerCategoria(InfoFileReport infoFileReport, Configurazione configurazione, List<FornitoreCensito> fornitoriDaAggiungere)
        {
            var worksheetName = infoFileReport.WorksheetName_ReportisticaPerTipologia; // REPORTISTICA PER TIPOLOGIA

            foreach (var fornitoreDaAggiungere in fornitoriDaAggiungere)
            {
                // Individuo la riga della categoria. Posso dare per scontato che la categoria del fornitore sia presente in elenco
                // in quanto la stessa viene precedentemente validata nel metodo 'VerificaCorrettezzaDeiNuoviFornitori'
                var rigaCorrente = infoFileReport.EPPlusHelper.GetFirstRowWithSpecificValue(
                    worksheetName: worksheetName,
                    rowToStartSearchFrom: configurazione.ReportisticaPerTipologia_PrimaRigaCategorieFornitori,
                    colToBeChecked: configurazione.ReportisticaPerTipologia_ColonnaCategorieFornitori,
                    valueToFind: fornitoreDaAggiungere.Categoria
                    );

                // Colonna J
                AccodaTestoAllaFormula(infoFileReport: infoFileReport,
                                        foglio: worksheetName,
                                        riga: rigaCorrente,
                                        colonna: configurazione.ReportisticaPerTipologia_Colonna_ConsumiSpeseAdOre_Ore,
                                        testoDaConcatenareAllaFormula: string.Format(configurazione.ReportisticaPerTipologia_Formula_ConsumiSpeseAdOre_Ore, fornitoreDaAggiungere.SiglaInReport));

                // Colonna K
                AccodaTestoAllaFormula(infoFileReport: infoFileReport,
                                        foglio: worksheetName,
                                        riga: rigaCorrente,
                                        colonna: configurazione.ReportisticaPerTipologia_Colonna_ConsumiSpeseAdOre_Euro,
                                        testoDaConcatenareAllaFormula: string.Format(configurazione.ReportisticaPerTipologia_Formula_ConsumiSpeseAdOre_Euro, fornitoreDaAggiungere.SiglaInReport));

                // Colonna L
                AccodaTestoAllaFormula(infoFileReport: infoFileReport,
                                        foglio: worksheetName,
                                        riga: rigaCorrente,
                                        colonna: configurazione.ReportisticaPerTipologia_Colonna_ConsumiSpeseLumpSum_Euro,
                                        testoDaConcatenareAllaFormula: string.Format(configurazione.ReportisticaPerTipologia_Formula_ConsumiSpeseLumpSum_Euro, fornitoreDaAggiungere.SiglaInReport));

                // Colonna N
                AccodaTestoAllaFormula(infoFileReport: infoFileReport,
                                        foglio: worksheetName,
                                        riga: rigaCorrente,
                                        colonna: configurazione.ReportisticaPerTipologia_Colonna_AllocateSpeseAdOre_Ore,
                                        testoDaConcatenareAllaFormula: string.Format(configurazione.ReportisticaPerTipologia_Formula_AllocateSpeseAdOre_Ore, fornitoreDaAggiungere.SiglaInReport));

                // Colonna O
                AccodaTestoAllaFormula(infoFileReport: infoFileReport,
                                        foglio: worksheetName,
                                        riga: rigaCorrente,
                                        colonna: configurazione.ReportisticaPerTipologia_Colonna_AllocateSpeseAdOre_Euro,
                                        testoDaConcatenareAllaFormula: string.Format(configurazione.ReportisticaPerTipologia_Formula_AllocateSpeseAdOre_Euro, fornitoreDaAggiungere.SiglaInReport));

                // Colonna P
                AccodaTestoAllaFormula(infoFileReport: infoFileReport,
                                        foglio: worksheetName,
                                        riga: rigaCorrente,
                                        colonna: configurazione.ReportisticaPerTipologia_Colonna_AllocateSpeseLumpSum_Euro,
                                        testoDaConcatenareAllaFormula: string.Format(configurazione.ReportisticaPerTipologia_Formula_AllocateSpeseLumpSum_Euro, fornitoreDaAggiungere.SiglaInReport));
            }
        }
        private void AccodaTestoAllaFormula(InfoFileReport infoFileReport, string foglio, int riga, int colonna, string testoDaConcatenareAllaFormula)
        {
            var currentFormula = infoFileReport.EPPlusHelper.GetFormula(foglio, riga, colonna);
            currentFormula = currentFormula.Replace("+IFERROR(", Environment.NewLine + "+IFERROR(");
            var newFormula = currentFormula + testoDaConcatenareAllaFormula;

            #region Verifica sulla lunghezza della nuova formula
            if (newFormula.Length >= Numbers.LIMITE_LUNGHEZZA_FORMULE_EXCEL)
            {
                throw new ManagedException(
                    tipologiaErrore: TipologiaErrori.DatoNonValido,
                    tipologiaCartella: TipologiaCartelle.ReportInput,
                    worksheetName: foglio,
                    rigaCella: riga,
                    colonnaCella: colonna,
                    nomeDatoErrore: NomiDatoErrore.Formula,
                    dato: null,
                    messaggioPerUtente: $"Impossibile procedere con l'aggiornamento della formula nella cella {(ColumnIDS)colonna}{riga} nel foglio '{foglio}'.\r\nLa sua lunghezza è già vicino al limite consentito da Excel. Per procedere è necessario semplificarla."
                );
            }
            #endregion

            infoFileReport.EPPlusHelper.SetFormula(foglio, riga, colonna, newFormula);
        }
    }
}