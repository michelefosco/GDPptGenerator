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
                var context = new StepContext(configurazione);
                context.SetContextFromInput(validateSourceFilesInput);
                var stepsSequence = new List<StepBase>
                {
                    new Step_Start_DebugInfoLogger(context),
                    new Step_ValidazioniPreliminari_InputFiles(context),
                    new Step_CreaLista_Applicablefilters(context),
                    new Step_CreaListe_Alias(context),
                    new Step_CreaLista_SildeToGenerate(context),
                    new Step_CreaLista_ItemsToExportAsImage(context),
                    new Step_TmpFolder_Pulizia(context),
                    new Step_EsitoFinale_Success(context)
                 };
                var esitoFinale = runStepSequence(stepsSequence, context);
                context.SettaEsitoFinale(esitoFinale);
                return new ValidateSourceFilesOutput(context);
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
                var context = new StepContext(configurazione);
                context.SetContextFromInput(buildPresentationInput);
                var stepsSequence = new List<StepBase>
                {
                    new Step_Start_DebugInfoLogger(context),
                    new Step_TmpFolder_Predisposizione(context),
                    new Step_BackupFile_DataSource(context),
                    new Step_CreaListe_Alias(context),
                    new Step_CreaLista_SildeToGenerate(context),
                    new Step_CreaLista_ItemsToExportAsImage(context),
                    new Step_CreaFilesImmaginiDaEsportare(context),
                    new Step_CreaFiles_Presentazioni(context),
                    new Step_TmpFolder_Pulizia(context),
                    new Step_EsitoFinale_Success(context)
                 };
                var esitoFinale = runStepSequence(stepsSequence, context);
                context.SettaEsitoFinale(esitoFinale);
                return new BuildPresentationOutput(context);
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

        private static EsitiFinali runStepSequence(List<StepBase> stepsSequence, StepContext context)
        {
            foreach (var step in stepsSequence)
            {
                var esitoStep = step.Do();

                if (esitoStep != EsitiFinali.Undefined)
                { return esitoStep; }
            }

            throw new Exception("The execution ended in an unexpected way.");
        }
    }
}
