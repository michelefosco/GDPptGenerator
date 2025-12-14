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
                    new Step_Start_Logger(stepContext),
                    new Step_Start_DebugInfoLogger(stepContext),

                    new Step_ValidazioniPreliminari_SourceFiles(stepContext),
                    new Step_ValidazioniPreliminari_SuperDettagli(stepContext),
                    new Step_CreaListe_Alias(stepContext),
                    new Step_CreaLista_Applicablefilters(stepContext),
                    new Step_CreaLista_SildeToGenerate(stepContext),
                    new Step_CreaLista_ItemsToExportAsImage(stepContext),

                    new Step_TmpFolder_Pulizia(stepContext),
                    new Step_Close_EPPlusHelpers(stepContext),

                    new Step_Stop_Logger(stepContext),
                    new Step_EsitoFinale_Success(stepContext)
                 };
                RunStepSequence(stepsSequence, stepContext);
                return new ValidateSourceFilesOutput(stepContext);
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


        #region Update Datasource and Build presentation
        public static UpdataDataSourceAndBuildPresentationOutput UpdataDataSourceAndBuildPresentation(UpdataDataSourceAndBuildPresentationInput buildPresentationInput)
        {
            var configurazione = ConfigurazioneHelper.GetConfigurazioneDefault();
            return updataDataSourceAndBuildPresentation(buildPresentationInput, configurazione);
        }

        public static UpdataDataSourceAndBuildPresentationOutput UpdataDataSourceAndBuildPresentation(UpdataDataSourceAndBuildPresentationInput buildPresentationInput, Configurazione configurazione)
        {
            if (configurazione == null)
            { throw new ArgumentNullException(nameof(configurazione)); }
            return updataDataSourceAndBuildPresentation(buildPresentationInput, configurazione);
        }

        private static UpdataDataSourceAndBuildPresentationOutput updataDataSourceAndBuildPresentation(UpdataDataSourceAndBuildPresentationInput buildPresentationInput, Configurazione configurazione)
        {
            try
            {
                var stepContext = new StepContext(configurazione);
                stepContext.SetContextFromInput(buildPresentationInput);
                var stepsSequence = new List<StepBase>
                {
                    new Step_Start_Logger(stepContext),
                    new Step_Start_DebugInfoLogger(stepContext),

                    new Step_VerificaEditabilita_DataSource_File(stepContext),
                    new Step_ValidazioniPreliminari_SourceFiles(stepContext),
                    new Step_ValidazioniPreliminari_SuperDettagli(stepContext),
                                       

                    #region Lettura info da DataSource
                    new Step_CreaListe_Alias(stepContext),
                    new Step_CreaLista_SildeToGenerate(stepContext),
                    new Step_CreaLista_ItemsToExportAsImage(stepContext),
                    #endregion

                    new Step_BackupFile_DataSource(stepContext),

                    #region Steps che modificano il file DataSource                  
                    new Step_DataSource_Editing_Start(stepContext),
                    new Step_ImportaDati_CN43N(stepContext),
                    new Step_ImportaDati_RunRate(stepContext),
                    new Step_ImportaDati_BudgetAndForecast(stepContext),
                    new Step_ImportaDati_SuperDettagli(stepContext),
                    new Step_ImpostaVarabiliInNameManager(stepContext),
                    new Step_AttivazioneOpzioneRefreshOnLoad(stepContext),
                    new Step_DataSource_Editing_Stop(stepContext),                    
                    #endregion

                    // fine operazioni fatte tramite libreria EPPlus
                    new Step_Close_EPPlusHelpers(stepContext),

                    #region Esportazione immagini da inserire nelle presentazioni
                    new Step_TmpFolder_Predisposizione(stepContext),
                    new Step_EsportaFileImmaginiDaExcel(stepContext),
                    #endregion

                    new Step_CreaFiles_Presentazioni(stepContext),
                    new Step_TmpFolder_Pulizia(stepContext),

                    new Step_Stop_Logger(stepContext),
                    new Step_EsitoFinale_Success(stepContext)
                 };
                RunStepSequence(stepsSequence, stepContext);
                return new UpdataDataSourceAndBuildPresentationOutput(stepContext);
            }
            catch (ManagedException managedException)
            {
                return new UpdataDataSourceAndBuildPresentationOutput(managedException);
            }
            catch (Exception ex)
            {
                return new UpdataDataSourceAndBuildPresentationOutput(new ManagedException(ex));
            }
        }
        #endregion


        private static void RunStepSequence(List<StepBase> stepsSequence, StepContext stepContext)
        {
            var startTime = DateTime.Now;
            foreach (var step in stepsSequence)
            {
                var esitoStep = step.DoStepTask();

                // Se ho un esito effettuo le ultime operazioni sullo StepContext e interrompo l'esecuzione
                if (esitoStep != EsitiFinali.Undefined)
                {
                    stepContext.ElapsedTime = DateTime.Now - startTime;
                    stepContext.SettaEsitoFinale(esitoStep);
                    stepContext.DebugInfoLogger.Beautify();
                    return;
                }
            }

            throw new Exception("The execution ended in an unexpected way.");
        }
    }
}