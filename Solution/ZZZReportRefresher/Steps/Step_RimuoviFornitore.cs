using ReportRefresher.Constants;
using ReportRefresher.Entities;
using ReportRefresher.Entities.Exceptions;
using ReportRefresher.Enums;
using ReportRefresher.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Rimozioni del fornitore indicato
    /// </summary>
    internal class Step_RimuoviFornitore : Step_Base
    {
        const string SIGLA_FORNITORE_ULTIMO = "ZZZZZZ";

        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            var siglaFornitore = context.Parameters["siglaFornitore"].ToString();

            verificaEsistenzaSiglaFornitore(context.FornitoriCensitiInReport, siglaFornitore);

            var categoriaFornitore = context.FornitoriCensitiInReport.First(_ => _.SiglaInReport.Equals(siglaFornitore, StringComparison.InvariantCultureIgnoreCase)).Categoria;
            var numeroReparti = context.RepartiCensitiInReport.Count;
            rimuoviFornitore(context.InfoFileReport, context.Configurazione, siglaFornitore, numeroReparti, categoriaFornitore);

            return null;
        }

        private void verificaEsistenzaSiglaFornitore(List<FornitoreCensito> fornitoriCensitiInReport, string siglaFornitore)
        {
            if (!fornitoriCensitiInReport.Any(_ => _.SiglaInReport.Equals(siglaFornitore, StringComparison.InvariantCultureIgnoreCase)))
            {
                throw new ManagedException(
                    tipologiaErrore: TipologiaErrori.DatoMancante,
                    tipologiaCartella: TipologiaCartelle.ReportInput,
                    nomeDatoErrore: NomiDatoErrore.SiglaFornitore,
                    //worksheetName: worksheetName,
                    //rigaCella: rigaCorrente,
                    //colonnaCella: configurazione.AnagraficaFornitori_ColonnaSigle,
                    dato: siglaFornitore
                    //percorsoFile: null
                    );
            }
        }

        private void rimuoviFornitore(InfoFileReport infoFileReport, Configurazione configurazione, string siglaFornitore, int numeroReparti, string catetoriaFornitore)
        {
            rimuoviFornitore_AnagraficaFornitori(infoFileReport, configurazione, siglaFornitore);
            rimuoviFornitore_ListaDati(infoFileReport, configurazione, siglaFornitore, numeroReparti);
            rimuoviFornitore_Sintesi(infoFileReport, configurazione, siglaFornitore);
            rimuoviFornitore_BudgetStudiIpotesi1(infoFileReport, configurazione, siglaFornitore);
            rimuoviFornitore_Reportistica(infoFileReport, configurazione, siglaFornitore);
            rimuoviFornitore_ReportisticaPerTipologia(infoFileReport, configurazione, siglaFornitore, catetoriaFornitore);
        }

        private void rimuoviFornitore_AnagraficaFornitori(InfoFileReport infoFileReport, Configurazione configurazione, string siglaFornitore)
        {
            var worksheetName = infoFileReport.WorksheetName_AnagraficaFornitori; // Anagrafica fornitori

            var rigaDaEliminare = infoFileReport.EPPlusHelper.GetFirstRowWithSpecificValue(
                worksheetName: worksheetName,
                rowToStartSearchFrom: configurazione.AnagraficaFornitori_PrimaRigaFornitori,
                colToBeChecked: configurazione.AnagraficaFornitori_ColonnaSigle,
                valueToFind: siglaFornitore);

            if (rigaDaEliminare < 0)
            {
                throw new ManagedException(
                    tipologiaErrore: TipologiaErrori.DatoMancante,
                    tipologiaCartella: TipologiaCartelle.ReportInput,
                    nomeDatoErrore: NomiDatoErrore.SiglaFornitore,
                    worksheetName: worksheetName,
                    dato: siglaFornitore
                    );
            }

            //elimino l'intera riga
            infoFileReport.EPPlusHelper.RemoveEntireRow(worksheetName, rigaDaEliminare);
        }

        private void rimuoviFornitore_ListaDati(InfoFileReport infoFileReport, Configurazione configurazione, string siglaFornitore, int numeroReparti)
        {
            var worksheetName = infoFileReport.WorksheetName_ListaDati; // Lista dati

            #region Marco la sigla del fornitore da eliminare con la costante che lo renderà ultimo nella lista della tabella SX
            var rigaDaEliminare = infoFileReport.EPPlusHelper.GetFirstRowWithSpecificValue(
                        worksheetName: worksheetName,
                        rowToStartSearchFrom: configurazione.ListaDati_PrimaRigaFornitori,
                        colToBeChecked: configurazione.ListaDati_ColonnaSiglaFornitoriTabellaSX,
                        valueToFind: siglaFornitore);

            if (rigaDaEliminare < 0)
            {
                throw new ManagedException(
                    tipologiaErrore: TipologiaErrori.DatoMancante,
                    tipologiaCartella: TipologiaCartelle.ReportInput,
                    nomeDatoErrore: NomiDatoErrore.SiglaFornitore,
                    worksheetName: worksheetName,
                    dato: siglaFornitore
                    );
            }

            infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaDaEliminare, configurazione.ListaDati_ColonnaSiglaFornitoriTabellaSX, SIGLA_FORNITORE_ULTIMO);
            #endregion


            #region Marco la sigla del fornitore da eliminare con la costante che lo renderà ultimo nella lista della tabella DX
            rigaDaEliminare = infoFileReport.EPPlusHelper.GetFirstRowWithSpecificValue(
                        worksheetName: worksheetName,
                        rowToStartSearchFrom: configurazione.ListaDati_PrimaRigaFornitori,
                        colToBeChecked: configurazione.ListaDati_ColonnaSiglaFornitoriTabellaDX,
                        valueToFind: siglaFornitore);

            if (rigaDaEliminare < 0)
            {
                throw new ManagedException(
                    tipologiaErrore: TipologiaErrori.DatoMancante,
                    tipologiaCartella: TipologiaCartelle.ReportInput,
                    nomeDatoErrore: NomiDatoErrore.SiglaFornitore,
                    worksheetName: worksheetName,
                    dato: siglaFornitore
                    );
            }

            infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaDaEliminare, configurazione.ListaDati_ColonnaSiglaFornitoriTabellaDX, SIGLA_FORNITORE_ULTIMO);
            #endregion


            // Ordino le tabelle
            OrdinaTabella_ListaDati(infoFileReport, configurazione);

            #region Svuoto contenuto riga marcata dalla tabella SX
            rigaDaEliminare = infoFileReport.EPPlusHelper.GetFirstRowWithSpecificValue(
                    worksheetName: worksheetName,
                    rowToStartSearchFrom: configurazione.ListaDati_PrimaRigaFornitori,
                    colToBeChecked: configurazione.ListaDati_ColonnaSiglaFornitoriTabellaSX,
                    valueToFind: SIGLA_FORNITORE_ULTIMO);

            infoFileReport.EPPlusHelper.CleanCellsContent(
                worksheetName: worksheetName,
                fromRow: rigaDaEliminare,
                fromCol: configurazione.ListaDati_ColonnaSiglaFornitoriTabellaSX,
                toRow: rigaDaEliminare,
                toCol: configurazione.ListaDati_ColonnaSiglaFornitoriTabellaSX + 2
                );
            #endregion

            #region Svuoto contenuto riga marcata dalla tabella DX
            rigaDaEliminare = infoFileReport.EPPlusHelper.GetFirstRowWithSpecificValue(
                worksheetName: worksheetName,
                rowToStartSearchFrom: configurazione.ListaDati_PrimaRigaFornitori,
                colToBeChecked: configurazione.ListaDati_ColonnaSiglaFornitoriTabellaDX,
                valueToFind: SIGLA_FORNITORE_ULTIMO);

            infoFileReport.EPPlusHelper.CleanCellsContent(
                worksheetName: worksheetName,
                fromRow: rigaDaEliminare,
                fromCol: configurazione.ListaDati_ColonnaSiglaFornitoriTabellaDX,
                toRow: rigaDaEliminare,
                toCol: configurazione.ListaDati_ColonnaSiglaFornitoriTabellaDX + numeroReparti
                );
            #endregion
        }

        private void rimuoviFornitore_Sintesi(InfoFileReport infoFileReport, Configurazione configurazione, string siglaFornitore)
        {
            var worksheetName = infoFileReport.WorksheetName_Sintesi; // Sintesi

            var ultimaRigaUsataNelFoglio = infoFileReport.EPPlusHelper.GetRowsLimit(worksheetName);
            var rigaCorrente = configurazione.Sintesi_Riga_PrimaConDati;

            // scorro tra tutte le righe della tabella
            while (rigaCorrente <= ultimaRigaUsataNelFoglio)
            {
                var siglaFornitoreRigaCorrente = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.Sintesi_ColonnaSiglaFornitore);

                if (siglaFornitoreRigaCorrente == null)
                {
                    break;
                }

                // elimino le righe contentente la sigla del fornitore da cancellare
                if (siglaFornitoreRigaCorrente.Equals(siglaFornitore, StringComparison.InvariantCultureIgnoreCase))
                {
                    infoFileReport.EPPlusHelper.RemoveEntireRow(worksheetName, rigaCorrente);
                    // avendo cancellato una riga, torno indietro di 1 con il conteggio
                    rigaCorrente--;
                }

                // vado alla riga successiva
                rigaCorrente++;
            }
        }

        private void rimuoviFornitore_BudgetStudiIpotesi1(InfoFileReport infoFileReport, Configurazione configurazione, string siglaFornitore)
        {
            var worksheetName = infoFileReport.WorksheetName_BudgetStudiIpotesi; // BUDGET STUDI -  IPOTESI 1

            var rigaDaEliminare = infoFileReport.EPPlusHelper.GetFirstRowWithSpecificValue(
                worksheetName: worksheetName,
                rowToStartSearchFrom: configurazione.IpotesiStudio_PrimaRigaFornitori,
                colToBeChecked: configurazione.IpotesiStudio_ColonnaSigle,
                valueToFind: siglaFornitore);

            if (rigaDaEliminare < 0)
            {
                throw new ManagedException(
                    tipologiaErrore: TipologiaErrori.DatoMancante,
                    tipologiaCartella: TipologiaCartelle.ReportInput,
                    nomeDatoErrore: NomiDatoErrore.SiglaFornitore,
                    worksheetName: worksheetName,
                    dato: siglaFornitore
                    );
            }

            //elimino l'intera riga
            infoFileReport.EPPlusHelper.RemoveEntireRow(worksheetName, rigaDaEliminare);
        }

        private void rimuoviFornitore_Reportistica(InfoFileReport infoFileReport, Configurazione configurazione, string siglaFornitore)
        {
            rimuoviFornitore_Reportistica_Tabella(infoFileReport, configurazione, siglaFornitore, configurazione.Reportistica_ColonnaSigle_SX);
            rimuoviFornitore_Reportistica_Tabella(infoFileReport, configurazione, siglaFornitore, configurazione.Reportistica_ColonnaSigle_DX);
        }

        private void rimuoviFornitore_Reportistica_Tabella(InfoFileReport infoFileReport, Configurazione configurazione, string siglaFornitore, int colonnaDelleSigle)
        {
            var worksheetName = infoFileReport.WorksheetName_Reportistica; // REPORTISTICA

            var ultimarigaUtilizzata = infoFileReport.EPPlusHelper.GetLastUsedRowForColumn(worksheetName, configurazione.Reportistica_PrimaRigaFornitori, colonnaDelleSigle);

            var rigaDaEliminare = infoFileReport.EPPlusHelper.GetFirstRowWithSpecificValue(
            worksheetName: worksheetName,
            rowToStartSearchFrom: configurazione.Reportistica_PrimaRigaFornitori,
            colToBeChecked: colonnaDelleSigle,
            valueToFind: siglaFornitore);

            if (rigaDaEliminare < 0)
            {
                throw new ManagedException(
                    tipologiaErrore: TipologiaErrori.DatoMancante,
                    tipologiaCartella: TipologiaCartelle.ReportInput,
                    nomeDatoErrore: NomiDatoErrore.SiglaFornitore,
                    worksheetName: worksheetName,
                    dato: siglaFornitore
                    );
            }

            // Marco la sigla del fornitore da eliminare con la costante che lo renderà ultimo nella lista
            infoFileReport.EPPlusHelper.SetValue(worksheetName, rigaDaEliminare, colonnaDelleSigle, SIGLA_FORNITORE_ULTIMO);

            // Ordino la tabella affinchè il fornitore da eliminare finisca in ultima posizione
            infoFileReport.EPPlusHelper.OrdinaTabella(
                 worksheetName: worksheetName,
                 fromRow: configurazione.Reportistica_PrimaRigaFornitori,
                 fromCol: colonnaDelleSigle,
                 toRow: ultimarigaUtilizzata,
                 toCol: colonnaDelleSigle
                 );

            rigaDaEliminare = infoFileReport.EPPlusHelper.GetFirstRowWithSpecificValue(
                worksheetName: worksheetName,
                rowToStartSearchFrom: configurazione.Reportistica_PrimaRigaFornitori,
                colToBeChecked: colonnaDelleSigle,
                valueToFind: SIGLA_FORNITORE_ULTIMO);


            // cancello il contenuto della riga del fornitore da cancellare
            infoFileReport.EPPlusHelper.CleanCellsContent(
                worksheetName: worksheetName,
                fromRow: rigaDaEliminare,
                fromCol: colonnaDelleSigle,
                toRow: rigaDaEliminare,
                toCol: colonnaDelleSigle
                );
        }

        private void rimuoviFornitore_ReportisticaPerTipologia(InfoFileReport infoFileReport, Configurazione configurazione, string siglaFornitore, string categoriaFornitore)
        {
            var worksheetName = infoFileReport.WorksheetName_ReportisticaPerTipologia; // REPORTISTICA PER TIPOLOGIA

            //// Individuo la riga della categoria. Posso dare per scontato che la categoria del fornitore sia presente in elenco
            //// in quanto la stessa viene precedentemente validata nel metodo 'VerificaCorrettezzaDeiNuoviFornitori'
            //var rigaDaEditare = infoFileReport.EPPlusHelper.GetFirstRowWithSpecificValue(
            //    worksheetName: worksheetName,
            //    rowToStartSearchFrom: configurazione.ReportisticaPerTipologia_PrimaRigaCategorieFornitori,
            //    colToBeChecked: configurazione.ReportisticaPerTipologia_ColonnaCategorieFornitori,
            //    valueToFind: categoriaFornitore
            //    );

            for (int rigaDaEditare = configurazione.ReportisticaPerTipologia_PrimaRigaCategorieFornitori; rigaDaEditare < 20; rigaDaEditare++)
            {
                // Colonna J
                rimuoviTestoDallaFormula(infoFileReport: infoFileReport,
                                        foglio: worksheetName,
                                        riga: rigaDaEditare,
                                        colonna: configurazione.ReportisticaPerTipologia_Colonna_ConsumiSpeseAdOre_Ore,
                                        testoDaRimuovereDallaFormula: string.Format(configurazione.ReportisticaPerTipologia_Formula_ConsumiSpeseAdOre_Ore, siglaFornitore));

                // Colonna K
                rimuoviTestoDallaFormula(infoFileReport: infoFileReport,
                                        foglio: worksheetName,
                                        riga: rigaDaEditare,
                                        colonna: configurazione.ReportisticaPerTipologia_Colonna_ConsumiSpeseAdOre_Euro,
                                        testoDaRimuovereDallaFormula: string.Format(configurazione.ReportisticaPerTipologia_Formula_ConsumiSpeseAdOre_Euro, siglaFornitore));

                // Colonna L
                rimuoviTestoDallaFormula(infoFileReport: infoFileReport,
                                        foglio: worksheetName,
                                        riga: rigaDaEditare,
                                        colonna: configurazione.ReportisticaPerTipologia_Colonna_ConsumiSpeseLumpSum_Euro,
                                        testoDaRimuovereDallaFormula: string.Format(configurazione.ReportisticaPerTipologia_Formula_ConsumiSpeseLumpSum_Euro, siglaFornitore));
                // Colonna N
                rimuoviTestoDallaFormula(infoFileReport: infoFileReport,
                                        foglio: worksheetName,
                                        riga: rigaDaEditare,
                                        colonna: configurazione.ReportisticaPerTipologia_Colonna_AllocateSpeseAdOre_Ore,
                                        testoDaRimuovereDallaFormula: string.Format(configurazione.ReportisticaPerTipologia_Formula_AllocateSpeseAdOre_Ore, siglaFornitore));

                // Colonna O
                rimuoviTestoDallaFormula(infoFileReport: infoFileReport,
                                        foglio: worksheetName,
                                        riga: rigaDaEditare,
                                        colonna: configurazione.ReportisticaPerTipologia_Colonna_AllocateSpeseAdOre_Euro,
                                        testoDaRimuovereDallaFormula: string.Format(configurazione.ReportisticaPerTipologia_Formula_AllocateSpeseAdOre_Euro, siglaFornitore));

                // Colonna P
                rimuoviTestoDallaFormula(infoFileReport: infoFileReport,
                                        foglio: worksheetName,
                                        riga: rigaDaEditare,
                                        colonna: configurazione.ReportisticaPerTipologia_Colonna_AllocateSpeseLumpSum_Euro,
                                        testoDaRimuovereDallaFormula: string.Format(configurazione.ReportisticaPerTipologia_Formula_AllocateSpeseLumpSum_Euro, siglaFornitore));
            }
        }

        private void rimuoviTestoDallaFormula(InfoFileReport infoFileReport, string foglio, int riga, int colonna, string testoDaRimuovereDallaFormula)
        {
            // leggo la formula attuale
            var currentFormula = infoFileReport.EPPlusHelper.GetFormula(foglio, riga, colonna);
            if (string.IsNullOrWhiteSpace(currentFormula))
            {
                return;
            }


            // adatto il testo da contatenare affinchè sia individuabile, le formule potrebbero avere diverse combinazioni di + e \n
            testoDaRimuovereDallaFormula = testoDaRimuovereDallaFormula.Replace(Environment.NewLine, "");
            testoDaRimuovereDallaFormula = testoDaRimuovereDallaFormula.Replace("+", "");

            if (!currentFormula.Contains(testoDaRimuovereDallaFormula))
            {
                return;
            }

            // rimuovo i NewLine dalla formula esistente
            var newFormula = currentFormula.Replace(Environment.NewLine, "");

            // rimuovo eventuali blank prima e dopo le porzioni di formula
            newFormula = rimuoviExtraBlankPrimaEDopoDiUnTesto(newFormula, "IFERROR"); //inizio della porzione di formula
            newFormula = rimuoviExtraBlankPrimaEDopoDiUnTesto(newFormula, ",0)"); // fine della porzione di formula

            // rimuovo il pezzo di formula del fornitore interessato
            newFormula = newFormula.Replace(testoDaRimuovereDallaFormula, "");

            // elimino gli eventuali doppi segni "+" che si sono creati nella formula
            newFormula = newFormula.Replace("++", "+");

            // ripristino i NewLine prima dei simboli "+" in modo da andare a capo in modo corretto
            newFormula = newFormula.Replace("+", Environment.NewLine + "+");

            // setto la nuova formula
            infoFileReport.EPPlusHelper.SetFormula(foglio, riga, colonna, newFormula);
        }

        private string rimuoviExtraBlankPrimaEDopoDiUnTesto(string formula, string testo)
        {
            for (int j = 1; j <= 5; j++)
            {
                formula = formula.Replace(" " + testo, testo);
                formula = formula.Replace(testo + " ", testo);
            }
            return formula;
        }
    }
}