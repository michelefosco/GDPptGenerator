using FilesEditor.Entities;
using FilesEditor.Helpers;
using FilesEditor.Steps;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilesEditor
{
    public class Editor
    {
        #region Metodi UpdateReports
        public static CreatePresentationsOutput CreatePresentations(CreatePresentationsInput createPresentationsInput)
        {
            var configurazione = ConfigurazioneHelper.GetConfigurazioneDefault();
            return updateReports(createPresentationsInput, configurazione);
        }

        public static CreatePresentationsOutput UpdateReports(CreatePresentationsInput createPresentationsInput, Configurazione configurazione)
        {
            if (configurazione == null)
            { throw new ArgumentNullException(nameof(configurazione)); }
            return updateReports(createPresentationsInput, configurazione);
        }

        private static CreatePresentationsOutput updateReports(CreatePresentationsInput createPresentationsInput, Configurazione configurazione)
        {
            var context = new StepContext(createPresentationsInput, configurazione);
            var stepsSequence = new List<Step_Base>
                {
                    new Step_PredisponiTmpFolder(context),
                    new Step_AggiornaDataSource(context),
                    new Step_CreaFilesImmagini(context),                    
                    new Step_CreaFilesPowerPoint(context),                    
                    //new Step_VerificaPercorsoNuovaVersioneFileReport(),
                    //new Step_Start_InfoFileController(),
                    //new Step_Start_InfoFileReport(),
                    //new Step_Lettura_CategorieFornitori(),
                    //new Step_Lettura_RepartiCensitiSuController(),
                    //new Step_Lettura_RepartiCensitiInReport(),
                    //new Step_Lettura_FornitoriCensitiInReport(),
                    //new Step_Lettura_CostiOrari_CategorieDeiFornitori(),
                    //new Step_Inserimento_NuoviFornitori_DaInterfaccia(),
                    //new Step_Lettura_Spese_NomiFornitoriNonCensiti_RigheSkippate(),
                    //new Step_Inserimento_NuoviFornitori_PresentiSoloInListaDati(),
                    //new Step_ValutazionePresenzaNuoviFornitori(),
                    //new Step_TabellaAvanzamento(),
                    //new Step_Tabella_ConsumiSpacchettati(),
                    //new Step_Tabella_Sintesi(),
                    //new Step_TabellaPrevisioneAfinire(),
                    //new Step_TabellaReportistica(),
                    //new Step_AttivazioneOpzioneRefreshOnLoad(),
                    //new Step_OrdinamentoTabelle(),
                    //new Step_Beautify(),
                    //new Step_SalvataggioNuovaVersioneFileReport(),
                    //new Step_ProduzioneContenutiExtraPerFileDebug(),
                    new Step_EsitoFinaleSuccess(context),
                 };
            return runStepSequence(stepsSequence, context);
        }
        #endregion

        private static CreatePresentationsOutput runStepSequence(List<Step_Base> stepsSequence, StepContext context)
        {
            foreach (var step in stepsSequence)
            {
                var stepResult = step.Do();

                if (stepResult != null)
                { return stepResult; }
                ;
            }

            throw new Exception("Gli step sono terminati in modo imprevisto");
        }

    }
}
