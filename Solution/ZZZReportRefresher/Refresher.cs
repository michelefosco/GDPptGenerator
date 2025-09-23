using ReportRefresher.Entities;
using ReportRefresher.Helpers;
using ReportRefresher.Steps;
using System;
using System.Collections.Generic;

namespace ReportRefresher
{
    public class Refresher
    {
        #region Metodi UpdateReports
        public static UpdateReportsOutput UpdateReports(UpdateReportsInput updateReportsInput)
        {
            var configurazione = ConfigurazioneHelper.GetConfigurazioneDefault();
            return updateReports(updateReportsInput, configurazione);
        }

        public static UpdateReportsOutput UpdateReports(UpdateReportsInput updateReportsInput, Configurazione configurazione)
        {
            if (configurazione == null)
            { throw new ArgumentNullException(nameof(configurazione)); }
            return updateReports(updateReportsInput, configurazione);
        }

        private static UpdateReportsOutput updateReports(UpdateReportsInput updateReportsInput, Configurazione configurazione)
        {
            var context = new StepContext(updateReportsInput, configurazione);
            var stepsSequence = new List<Step_Base>
                {
                    new Step_Start_FileDebugHelper(),
                    new Step_VerificaPercorsoNuovaVersioneFileReport(),
                    new Step_Start_InfoFileController(),
                    new Step_Start_InfoFileReport(),
                    new Step_Lettura_CategorieFornitori(),
                    new Step_Lettura_RepartiCensitiSuController(),
                    new Step_Lettura_RepartiCensitiInReport(),
                    new Step_Lettura_FornitoriCensitiInReport(),
                    new Step_Lettura_CostiOrari_CategorieDeiFornitori(),
                    new Step_Inserimento_NuoviFornitori_DaInterfaccia(),
                    new Step_Lettura_Spese_NomiFornitoriNonCensiti_RigheSkippate(),
                    new Step_Inserimento_NuoviFornitori_PresentiSoloInListaDati(),
                    new Step_ValutazionePresenzaNuoviFornitori(),
                    new Step_TabellaAvanzamento(),
                    new Step_Tabella_ConsumiSpacchettati(),
                    new Step_Tabella_Sintesi(),
                    new Step_TabellaPrevisioneAfinire(),
                    new Step_TabellaReportistica(),
                    new Step_AttivazioneOpzioneRefreshOnLoad(),
                    new Step_OrdinamentoTabelle(),
                    new Step_Beautify(),
                    new Step_SalvataggioNuovaVersioneFileReport(),
                    new Step_ProduzioneContenutiExtraPerFileDebug(),
                    new Step_EsitoFinaleSuccess(),
                 };
            return runStepSequence(stepsSequence, context);
        }
        #endregion

        #region Metodi Gestione fornitori
        public static UpdateReportsOutput AddFornitori(UpdateReportsInput updateReportsInput)
        {
            var configurazione = ConfigurazioneHelper.GetConfigurazioneDefault();
            var context = new StepContext(updateReportsInput, configurazione);
            var stepsSequence = new List<Step_Base>
                {
                    new Step_Start_InfoFileReport(),
                    new Step_Lettura_CategorieFornitori(),
                    new Step_Lettura_RepartiCensitiInReport(),
                    new Step_Lettura_FornitoriCensitiInReport(), // questo step mi assicura che la tabella di "Anagrafica fornitori" sia ben formata
                    new Step_Lettura_CostiOrari_CategorieDeiFornitori(),
                    new Step_Inserimento_NuoviFornitori_DaInterfaccia(),
                    new Step_OrdinamentoTabelle(),
                    new Step_Beautify(),
                    new Step_SalvataggioNuovaVersioneFileReport(),
                    new Step_EsitoFinaleSuccess(),
                 };
            return runStepSequence(stepsSequence, context);
        }
        public static UpdateReportsOutput DeleteFornitore(string fileReport_FilePath, string newReport_FilePath, string siglaFornitore)
        {
            if (string.IsNullOrWhiteSpace(fileReport_FilePath)) { throw new ArgumentNullException(nameof(fileReport_FilePath)); }
            if (string.IsNullOrWhiteSpace(newReport_FilePath)) { throw new ArgumentNullException(nameof(newReport_FilePath)); }            
            if (string.IsNullOrWhiteSpace(siglaFornitore)) { throw new ArgumentNullException(nameof(siglaFornitore)); }

            var context = buildStepContextToUpdateFileReport(fileReport_FilePath, newReport_FilePath);
            context.Parameters.Add("siglaFornitore", siglaFornitore);
            var stepsSequence = new List<Step_Base>
                {
                    new Step_Start_InfoFileReport(),
                    new Step_Lettura_CategorieFornitori(),
                    new Step_Lettura_RepartiCensitiInReport(), // necessario per avere le info sul numero di reparti presenti nel file report
                    new Step_Lettura_FornitoriCensitiInReport(), // questo step mi assicura che la tabella di "Anagrafica fornitori" sia ben formata
                    new Step_Lettura_CostiOrari_CategorieDeiFornitori(),
                    new Step_RimuoviFornitore(),
                    new Step_OrdinamentoTabelle(),  // necessario per ordinare e quindi compattare le tabelle in cui sono state azzerate le celle (anzichè rimuovere l'intera riga)
                    new Step_SalvataggioNuovaVersioneFileReport(),
                    new Step_EsitoFinaleSuccess(),
                 };
            return runStepSequence(stepsSequence, context);
        }
        public static UpdateReportsOutput UpdateFornitore(string fileReport_FilePath, string newReport_FilePath, string siglaFornitore, string nuovoNomeFornitore)
        {
            if (string.IsNullOrWhiteSpace(fileReport_FilePath)) { throw new ArgumentNullException(nameof(fileReport_FilePath)); }
            if (string.IsNullOrWhiteSpace(newReport_FilePath)) { throw new ArgumentNullException(nameof(newReport_FilePath)); }
            if (string.IsNullOrWhiteSpace(siglaFornitore)) { throw new ArgumentNullException(nameof(siglaFornitore)); }
            if (string.IsNullOrWhiteSpace(nuovoNomeFornitore)) { throw new ArgumentNullException(nameof(nuovoNomeFornitore)); }

            var context = buildStepContextToUpdateFileReport(fileReport_FilePath, newReport_FilePath);
            context.Parameters.Add("siglaFornitore", siglaFornitore);
            context.Parameters.Add("nuovoNomeFornitore", nuovoNomeFornitore);
            var stepsSequence = new List<Step_Base>
                {
                    new Step_Start_InfoFileReport(),
                    new Step_Lettura_FornitoriCensitiInReport(), // questo step mi assicura che la tabella di "Anagrafica fornitori" sia ben formata
                    new Step_AggiornaNomeFornitore(),
                    new Step_SalvataggioNuovaVersioneFileReport(),
                    new Step_EsitoFinaleSuccess(),
                 };
            return runStepSequence(stepsSequence, context);
        }

        public static UpdateReportsOutput AddFornitore(string fileReport_FilePath, string newReport_FilePath, FornitoreCensito nuovoFornitore)
        {
            UpdateReportsInput input = new UpdateReportsInput(DateTime.MinValue, string.Empty, fileReport_FilePath, newReport_FilePath);
            input.SettaNuoviFornitori(new List<FornitoreCensito> { nuovoFornitore });
            return Refresher.AddFornitori(input);
        }
        #endregion

        #region Metodi info
        public static UpdateReportsOutput IsaValidReportFile(string fileReport_FilePath)
        {
            var context = buildStepContextToUpdateFileReport(fileReport_FilePath, null);
            var stepsSequence = new List<Step_Base>
                {
                    new Step_Start_InfoFileReport(),
                    new Step_EsitoFinaleSuccess(),
                 };
            return runStepSequence(stepsSequence, context);
        }
        public static UpdateReportsOutput GetInfoFromFile(string fileReport_FilePath)
        {
            var context = buildStepContextToUpdateFileReport(fileReport_FilePath, null);
            var stepsSequence = new List<Step_Base>
                {
                    new Step_Start_InfoFileReport(),
                    new Step_Lettura_CategorieFornitori(),
                    new Step_Lettura_RepartiCensitiInReport(),
                    new Step_Lettura_FornitoriCensitiInReport(),
                    new Step_Lettura_CostiOrari_CategorieDeiFornitori(),
                    new Step_EsitoFinaleSuccess(),
                 };

            return runStepSequence(stepsSequence, context);
        }
        #endregion

        private static UpdateReportsOutput runStepSequence(List<Step_Base> stepsSequence, StepContext context)
        {
            foreach (var step in stepsSequence)
            {
                var stepResult = step.Do(context);

                if (stepResult != null)
                { return stepResult; };
            }

            throw new Exception("Gli step sono terminati in modo imprevisto");
        }

        private static StepContext buildStepContextToUpdateFileReport(string fileReport_FilePath, string newReport_FilePath)
        {
            var updateReportsInput = new UpdateReportsInput(
                    dataAggiornamento: DateTime.Today,
                    fileController_FilePath: null,
                    fileReport_FilePath: fileReport_FilePath,
                    newReport_FilePath: newReport_FilePath,
                    fileDebug_FilePath: null
                );
            var configurazione = ConfigurazioneHelper.GetConfigurazioneDefault();
            return new StepContext(updateReportsInput, configurazione);
        }
    }
}