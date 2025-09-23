using PptGenerator.Entities;
using PptGenerator.Entities.Exceptions;
using PptGenerator.Enums;

namespace PptGenerator.Steps
{
    //todo: rendere internal ma accessibile ai test
    abstract public class Step_Base
    {
        public Step_Base()
        { }

        internal abstract UpdateReportsOutput DoSpecificTask(StepContext context);

        internal UpdateReportsOutput Do(StepContext context)
        {
            try
            {
                return DoSpecificTask(context);
            }
            catch (ManagedException ex)
            {
                context.UpdateReportsOutput.SettaManagedException(ex);
                context.UpdateReportsOutput.SettaEsitoFinale(EsitiFinali.Failure);

                #region Se l'opzione è attiva vengono Evidenziati gli errori direttamente nei file di input
                if (context.UpdateReportsInput.EvidenziaErroriNelFileDiInput)
                {
                    // vengono esclusi gli errori per fogli mancante, in quanto non sarebbe possibile selezionarlo
                    if (!string.IsNullOrEmpty(ex.WorksheetName) && ex.TipologiaErrore != TipologiaErrori.FoglioMancante)
                    {
                        switch (ex.TipologiaCartella)
                        {
                            //case TipologiaCartelle.ReportInput:
                            //    context.InfoFileReport.EPPlusHelper.SetErrorCell(ex.WorksheetName, ex.RigaCella, ex.ColonnaCella);
                            //    context.InfoFileReport.EPPlusHelper.Save();
                            //    break;
                            //case TipologiaCartelle.Controller:
                            //    context.InfoFileController.EPPlusHelper.SetErrorCell(ex.WorksheetName, ex.RigaCella, ex.ColonnaCella);
                            //    context.InfoFileController.EPPlusHelper.Save();
                            //    break;
                        }
                    }
                }
                #endregion

                ChiusuraFileDebug(context);

                return context.UpdateReportsOutput;
            }
        }

        internal void ChiusuraFileDebug(StepContext context)
        {
            if (context.DebugInfoLogger != null)
            {
                context.DebugInfoLogger.LogUpdateReportsInput(context.UpdateReportsInput);
                context.DebugInfoLogger.LogUpdateReportsOutput(context.UpdateReportsOutput);
                context.DebugInfoLogger.Beautify();
                context.DebugInfoLogger.Save();
            }
        }

        //internal void OrdinaTabella_ListaDati(InfoFileReport infoFileReport, Configurazione configurazione)
        //{
        //    var worksheetName = infoFileReport.WorksheetName_ListaDati; // Lista dati

        //    #region lettura dei reparti dalla tabella dei costi specifici per reparto
        //    var rigaReparti = configurazione.ListaDati_RigaReparti;

        //    var colonnaCorrente = configurazione.ListaDati_PrimaColonnaReparti;
        //    while (true)
        //    {
        //        var nomeReparto = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaReparti, colonnaCorrente);
        //        if (string.IsNullOrWhiteSpace(nomeReparto)) // mi interrompo con l'ultima cella con eventuale valore blank
        //        { break; }
        //        colonnaCorrente++;
        //    }
        //    var ultimaColonnaReparti = colonnaCorrente - 1;
        //    #endregion

        //    var ultimarigaUtilizzata = infoFileReport.EPPlusHelper.GetLastUsedRowForColumn(worksheetName, configurazione.ListaDati_PrimaRigaFornitori, configurazione.ListaDati_ColonnaSiglaFornitoriTabellaSX);
        //    if (ultimarigaUtilizzata <= configurazione.ListaDati_PrimaRigaFornitori)
        //    { return; /* meno di 2 righe presenti, non è neccessario ordinare*/}

        //    VerificaAssenzaFormuleNellaColonnaSigle(
        //        infoFileReport: infoFileReport,
        //        worksheetName: worksheetName,
        //        colToBeChecked: configurazione.ListaDati_ColonnaSiglaFornitoriTabellaSX,
        //        rowFrom: configurazione.ListaDati_PrimaRigaFornitori,
        //        rowTo: ultimarigaUtilizzata
        //        );

        //    // Ordina la tabella SX per "Sigla"
        //    infoFileReport.EPPlusHelper.OrdinaTabella(
        //        worksheetName: worksheetName,
        //        fromRow: configurazione.ListaDati_PrimaRigaFornitori,
        //        fromCol: configurazione.ListaDati_ColonnaSiglaFornitoriTabellaSX,
        //        toRow: ultimarigaUtilizzata,
        //        toCol: configurazione.ListaDati_ColonnaCategoriaFornitori
        //        );

        //    VerificaAssenzaFormuleNellaColonnaSigle(
        //        infoFileReport: infoFileReport,
        //        worksheetName: worksheetName,
        //        colToBeChecked: configurazione.ListaDati_ColonnaSiglaFornitoriTabellaDX,
        //        rowFrom: configurazione.ListaDati_PrimaRigaFornitori,
        //        rowTo: ultimarigaUtilizzata
        //        );

        //    // Ordina la tabella DX per "Sigla"
        //    infoFileReport.EPPlusHelper.OrdinaTabella(
        //        worksheetName: worksheetName,
        //        fromRow: configurazione.ListaDati_PrimaRigaFornitori,
        //        fromCol: configurazione.ListaDati_ColonnaSiglaFornitoriTabellaDX,
        //        toRow: ultimarigaUtilizzata,
        //        toCol: ultimaColonnaReparti
        //        );
        //}

        //internal void VerificaAssenzaFormuleNellaColonnaSigle(InfoFileReport infoFileReport, string worksheetName, int colToBeChecked, int rowFrom, int rowTo)
        //{
        //    for (int rowToBeChecked = rowFrom; rowToBeChecked <= rowTo; rowToBeChecked++)
        //    {
        //        if (infoFileReport.EPPlusHelper.IsFormula(worksheetName, rowToBeChecked, colToBeChecked))
        //        {
        //            var foundFormula = "=" + infoFileReport.EPPlusHelper.GetFormula(worksheetName, rowToBeChecked, colToBeChecked);
        //            throw new ManagedException(
        //                 tipologiaErrore: TipologiaErrori.DatoNonValido,
        //                 tipologiaCartella: TipologiaCartelle.ReportInput,
        //                 nomeDatoErrore: NomiDatoErrore.SiglaFornitore,
        //                 messaggioPerUtente: "La cella contiente una formula inattesa. Dovrebbe contenere un valore.",
        //                 worksheetName: worksheetName,
        //                 rigaCella: rowToBeChecked,
        //                 colonnaCella: colToBeChecked,
        //                 dato: foundFormula,
        //                 percorsoFile: null
        //                 );
        //        }
        //    }
        //}
    }
}