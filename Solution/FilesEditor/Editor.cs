using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using FilesEditor.Steps;
using FilesEditor.Steps.BuildPresentation;
using FilesEditor.Steps.ValidateSourceFiles;
using System;
using System.Collections.Generic;

namespace FilesEditor
{
    public class Editor
    {
        #region ValidateSourceFiles
        public static ValidateSourceFilesOutput ValidateSourceFiles(ValidateSourceFilesInput validateSourceFilesInput)
        {
            var configurazione = ConfigurazioneHelper.GetConfigurazioneDefault();
            return validateSourceFiles(validateSourceFilesInput, configurazione);
        }

        public static ValidateSourceFilesOutput ValidateSourceFiles(ValidateSourceFilesInput validateSourceFilesInput, Configurazione configurazione)
        {
            if (configurazione == null)
            { throw new ArgumentNullException(nameof(configurazione)); }
            return validateSourceFiles(validateSourceFilesInput, configurazione);
        }

        private static ValidateSourceFilesOutput validateSourceFiles(ValidateSourceFilesInput validateSourceFilesInput, Configurazione configurazione)
        {
            try
            {
                var stepContext = new StepContext(configurazione);
                stepContext.SetContextFromInput(validateSourceFilesInput);
                var stepsSequence = new List<StepBase>
                {
                    new Step_Start_DebugInfoLogger(stepContext),
                    new Step_ValidazioniPreliminari_SourceFiles(stepContext),
                    new Step_CreaListe_Alias(stepContext),
                    new Step_CreaLista_Applicablefilters(stepContext),
                    new Step_CreaLista_SildeToGenerate(stepContext),
                    new Step_CreaLista_ItemsToExportAsImage(stepContext),
                    new Step_TmpFolder_Pulizia(stepContext),
                    new Step_EsitoFinale_Success(stepContext)
                 };
                RunStepSequence(stepsSequence, stepContext);
                return new ValidateSourceFilesOutput(stepContext);
                //var esitoFinale = RunStepSequence(stepsSequence, stepContext);
                //stepContext.SettaEsitoFinale(esitoFinale);

                //stepContext.DebugInfoLogger.Beautify();
                //return new ValidateSourceFilesOutput(stepContext);
            }
            catch (ManagedException managedException)
            {
                return new ValidateSourceFilesOutput(managedException);
            }
            catch (Exception ex)
            {
                return new ValidateSourceFilesOutput(new ManagedException(ex));
            }
        }
        #endregion


        #region Build presentation
        public static BuildPresentationOutput BuildPresentation(BuildPresentationInput buildPresentationInput)
        {
            var configurazione = ConfigurazioneHelper.GetConfigurazioneDefault();
            return buildPresentation(buildPresentationInput, configurazione);
        }

        public static BuildPresentationOutput BuildPresentation(BuildPresentationInput buildPresentationInput, Configurazione configurazione)
        {
            if (configurazione == null)
            { throw new ArgumentNullException(nameof(configurazione)); }
            return buildPresentation(buildPresentationInput, configurazione);
        }

        private static BuildPresentationOutput buildPresentation(BuildPresentationInput buildPresentationInput, Configurazione configurazione)
        {
            try
            {
                var stepContext = new StepContext(configurazione);
                stepContext.SetContextFromInput(buildPresentationInput);
                var stepsSequence = new List<StepBase>
                {
                    new Step_Start_DebugInfoLogger(stepContext),
                    new Step_ValidazioniPreliminari_SourceFiles(stepContext),
                    new Step_VerificaEditabilita_DataSource_File(stepContext),
                    new Step_TmpFolder_Predisposizione(stepContext),
                    new Step_BackupFile_DataSource(stepContext),
                    new Step_CreaListe_Alias(stepContext),
                    new Step_CreaLista_SildeToGenerate(stepContext),
                    new Step_CreaLista_ItemsToExportAsImage(stepContext),
                    #region Steps che modificano il file DataSource - Inizio
                    new Step_ImportaDati_RunRate(stepContext),
                    new Step_ImportaDati_BudgetAndForecast(stepContext),
                    //new Step_ImportaDati_SuperDettagli(context),
                    new Step_ImportaDati_SuperDettagli_2(stepContext),
                    new Step_ImpostaVarabiliInNameManager(stepContext),
                    new Step_AttivazioneOpzioneRefreshOnLoad(stepContext),
                    new Step_SalvaFile_DataSource(stepContext),
                    #endregion
                    new Step_EsportaFileImmaginiDaExcel(stepContext),
                    new Step_CreaFiles_Presentazioni(stepContext),
                    new Step_TmpFolder_Pulizia(stepContext),
                    new Step_EsitoFinale_Success(stepContext)
                 };
                RunStepSequence(stepsSequence, stepContext);
                return new BuildPresentationOutput(stepContext);
            }
            catch (ManagedException managedException)
            {
                return new BuildPresentationOutput(managedException);
            }
            catch (Exception ex)
            {
                return new BuildPresentationOutput(new ManagedException(ex));
            }
        }
        #endregion

        private static void RunStepSequence(List<StepBase> stepsSequence, StepContext stepContext)
        {
            foreach (var step in stepsSequence)
            {
                var esitoStep = step.DoStepTask();

                // Se ho un esito effettuo le ultime operazioni sullo StepContext e interrompo l'esecuzione
                if (esitoStep != EsitiFinali.Undefined)
                {
                    stepContext.SettaEsitoFinale(esitoStep);
                    stepContext.DebugInfoLogger.Beautify();
                    return;
                }
            }

            throw new Exception("The execution ended in an unexpected way.");
        }
    }
}
