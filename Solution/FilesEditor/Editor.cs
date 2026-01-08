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
            #region Preparazione StepContext
            var stepContext = new StepContext(configurazione);
            stepContext.SetContextFromInput(validateSourceFilesInput);
            #endregion

            #region Preparazione Steps Sequence
            var stepsSequence = new List<StepBase>
                {
                    new Step_Start_Logger(stepContext),

                    new Step_ValidazioniPreliminari_SourceFiles(stepContext),
                    new Step_ValidazioniPreliminari_SuperDettagli(stepContext),
                    new Step_CreaListe_Alias(stepContext),
                    new Step_CreaLista_Applicablefilters(stepContext),
                    new Step_CreaLista_SildeToGenerate(stepContext),
                    new Step_CreaLista_ItemsToExportAsImage(stepContext),

                    new Step_TmpFolder_Pulizia(stepContext),
                    new Step_Close_EPPlusHelpers(stepContext),

                    new Step_EsitoFinale_Success(stepContext)
                 };
            #endregion

            try
            {

                RunStepSequence(stepsSequence, stepContext);
                return new ValidateSourceFilesOutput(stepContext);
            }
            catch (ManagedException managedException)
            {
                return new ValidateSourceFilesOutput(stepContext, managedException);
            }
            catch (Exception ex)
            {
                return new ValidateSourceFilesOutput(stepContext, new ManagedException(ex));
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
            #region Preparazione StepContext
            var stepContext = new StepContext(configurazione);
            stepContext.SetContextFromInput(buildPresentationInput);
            if (buildPresentationInput.BuildPresentationOnly)
            {
                stepContext.SetDatasourceStatus_ImportDatiCompletato();
            }
            #endregion


            #region Preparazione Steps Sequence
            var stepsSequence = new List<StepBase> { new Step_Start_Logger(stepContext) };

            if (!buildPresentationInput.BuildPresentationOnly)
            {
                // Verifica che il file DataSource sia apribile in modalità editabile
                stepsSequence.Add(new Step_VerificaEditabilita_DataSource_File(stepContext));

                // Validazioni preliminari sui file di origine
                stepsSequence.Add(new Step_ValidazioniPreliminari_SourceFiles(stepContext));
                stepsSequence.Add(new Step_ValidazioniPreliminari_SuperDettagli(stepContext));

                // Lettura alias necessario per le importazioni
                stepsSequence.Add(new Step_CreaListe_Alias(stepContext));
            }

            // Lettura info da DataSource
            stepsSequence.Add(new Step_CreaLista_SildeToGenerate(stepContext));
            stepsSequence.Add(new Step_CreaLista_ItemsToExportAsImage(stepContext));
            

            if (!buildPresentationInput.BuildPresentationOnly)
            {
                // Backup del file DataSource
                stepsSequence.Add(new Step_BackupFile_DataSource(stepContext));

               // Steps che modificano il file DataSource                  
                stepsSequence.Add(new Step_DataSource_Editing_Start(stepContext));
                stepsSequence.Add(new Step_ImportaDati_CN43N(stepContext));
                stepsSequence.Add(new Step_ImportaDati_RunRate(stepContext));
                stepsSequence.Add(new Step_ImportaDati_BudgetAndForecast(stepContext));
                stepsSequence.Add(new Step_ImportaDati_SuperDettagli(stepContext));
                stepsSequence.Add(new Step_ImpostaVarabiliInNameManager(stepContext));
                stepsSequence.Add(new Step_AttivazioneOpzioneRefreshOnLoad(stepContext));
                stepsSequence.Add(new Step_DataSource_Stop_Editing(stepContext));
            }

            // fine operazioni fatte tramite libreria EPPlus
            stepsSequence.Add(new Step_Close_EPPlusHelpers(stepContext));

            // Esportazione immagini da inserire nelle presentazioni
            stepsSequence.Add(new Step_TmpFolder_Predisposizione(stepContext));
            stepsSequence.Add(new Step_EsportaFileImmaginiDaExcel(stepContext));

            // Creazione file presentazioni
            stepsSequence.Add(new Step_CreaFiles_Presentazioni(stepContext));

            // Pulizia cartella temporanea
            stepsSequence.Add(new Step_TmpFolder_Pulizia(stepContext));

            // Esito finale positivo
            stepsSequence.Add(new Step_EsitoFinale_Success(stepContext));
            #endregion


            #region Esecuzione Steps Sequence
            try
            {
                RunStepSequence(stepsSequence, stepContext);
                return new UpdataDataSourceAndBuildPresentationOutput(stepContext);
            }
            catch (ManagedException managedException)
            {
                return new UpdataDataSourceAndBuildPresentationOutput(stepContext, managedException);
            }
            catch (Exception ex)
            {
                return new UpdataDataSourceAndBuildPresentationOutput(stepContext, new ManagedException(ex));
            }
            #endregion
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
                    return;
                }
            }

            throw new Exception("The execution ended in an unexpected way.");
        }
    }
}