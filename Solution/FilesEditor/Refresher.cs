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
                    //new Step_Start_FileDebugHelper(),
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
                    //new Step_EsitoFinaleSuccess(),
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
                { return stepResult; }
                ;
            }

            throw new Exception("Gli step sono terminati in modo imprevisto");
        }

    }
}
