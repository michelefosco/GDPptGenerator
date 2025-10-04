using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using FilesEditor.Steps;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilesEditor
{
    public class Editor
    {
        #region Metodi pubblici


        public static CreatePresentationsOutput CreatePresentations(CreatePresentationsInput createPresentationsInput)
        {
            var configurazione = ConfigurazioneHelper.GetConfigurazioneDefault();
            return createPresentations(createPresentationsInput, configurazione);
        }

        public static CreatePresentationsOutput CreatePresentations(CreatePresentationsInput createPresentationsInput, Configurazione configurazione)
        {
            if (configurazione == null)
            { throw new ArgumentNullException(nameof(configurazione)); }
            return createPresentations(createPresentationsInput, configurazione);
        }

        private static CreatePresentationsOutput createPresentations(CreatePresentationsInput createPresentationsInput, Configurazione configurazione)
        {
            var context = new StepContext(createPresentationsInput, configurazione);
            var stepsSequence = new List<Step_Base>
                {
                    new Step_PredisponiTmpFolder(context),
                    new Step_Start_FileDebugHelper(context),
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


        public static ValidaSourceFilesOutput ValidaSourceFiles(ValidaSourceFilesInput validaSourceFilesInput)
        {
            // Lettura Opzione dal datasource

            var dataSourceTemplateFile = Path.Combine(validaSourceFilesInput.TemplatesFolder, Constants.FileNames.DATA_SOURCE_FILENAME);
            var opzioniUtente = GetOpzioniUtente(dataSourceTemplateFile);

            // lettura info da 1° file

            // lettura info da 2° file

            // lettura info da 3° file

            // lettura info da 4° file

            // verifica applicabilità filtri

            // completa Valori sele

            return new ValidaSourceFilesOutput
            {
                OpzioniUtente = opzioniUtente
            };
        }


        private static OpzioniUtente GetOpzioniUtente(string filePath)
        {
            EPPlusExtensions.EPPlusHelper ePPlusHelper = GetHelperForExistingFile(filePath, TipologiaCartelle.DataSource_Template);

            //todo leggere da file
            var filtriPossibili = new List<FilterItems>();
            filtriPossibili.Add(new FilterItems
            {
                Tabella = "SUPERDETTAGLI",
                Campo = "ProjType Cluster 2_",
                ValoriPossibili = new List<string>(),
                ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED }
            });
            filtriPossibili.Add(new FilterItems { Tabella = "SUPERDETTAGLI", Campo = "BussinessArea Cluster 1_", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "SUPERDETTAGLI", Campo = "BusinessArea_", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "SUPERDETTAGLI", Campo = "ProjType_", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "FORECAST", Campo = "Proj type cluster 2", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "FORECAST", Campo = "Business", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "BUDGET", Campo = "EngUnit area cluster 1_", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });
            filtriPossibili.Add(new FilterItems { Tabella = "BUDGET", Campo = "CATEGORIA_", ValoriSelezionati = new List<string> { Values.ALLFILTERSAPPLIED } });

            // Implement your logic to read user options from the specified file
            return new OpzioniUtente
            {
                FiltriPossibili = filtriPossibili
            };
        }

        public static bool IsBudgetFileOk(string filePath)
        {
            // check file existence

            // check expected worksheet names

            // check expected columns in each worksheet

            // Implement your logic to check if the budget file is OK
            return true;
        }


        static internal EPPlusHelper GetHelperForExistingFile(string filePath, TipologiaCartelle tipologiaCartella)
        {
            var ePPlusHelper = new EPPlusHelper();
            if (!ePPlusHelper.Open(filePath))
            {
                throw new ManagedException(
                    tipologiaErrore: TipologiaErrori.UnableToOpenFile,
                    tipologiaCartella: tipologiaCartella,
                    messaggioPerUtente: MessaggiErrorePerUtente.UnableToOpenFile,
                    percorsoFile: filePath);
            }
            return ePPlusHelper;
        }


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
