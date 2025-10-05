using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using FilesEditor.Steps;
using FilesEditor.Steps.CreatePresentation;
using System;
using System.Collections.Generic;
using System.IO;

namespace FilesEditor
{
    public class Editor
    {
        #region CreatePresentations
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
                    new Step_EsitoFinaleSuccess(context),
                 };
            return createPresentationsRunStepSequence(stepsSequence, context);
        }

        private static CreatePresentationsOutput createPresentationsRunStepSequence(List<Step_Base> stepsSequence, StepContext context)
        {
            foreach (var step in stepsSequence)
            {
                var stepResult = step.Do();

                if (stepResult != null)
                { return stepResult; }
            }
            throw new Exception("The execution ended in an unexpected way.");
        }
        #endregion


        #region ValidaSourceFiles
        public static ValidaSourceFilesOutput ValidaSourceFiles(ValidaSourceFilesInput validaSourceFilesInput)
        {
            var configurazione = ConfigurazioneHelper.GetConfigurazioneDefault();
            return validaSourceFiles(validaSourceFilesInput, configurazione);
        }

        public static ValidaSourceFilesOutput ValidaSourceFiles(ValidaSourceFilesInput validaSourceFilesInput, Configurazione configurazione)
        {
            if (configurazione == null)
            { throw new ArgumentNullException(nameof(configurazione)); }
            return validaSourceFiles(validaSourceFilesInput, configurazione);
        }

        private static ValidaSourceFilesOutput validaSourceFiles(ValidaSourceFilesInput validaSourceFilesInput, Configurazione configurazione)
        {
            // Lettura Opzione dal datasource

            var dataSourceTemplateFile = Path.Combine(validaSourceFilesInput.TemplatesFolder, Constants.FileNames.DATA_SOURCE_TEMPLATE_FILENAME);
            var ePPlusHelper = GetHelperForExistingFile(dataSourceTemplateFile, FileTypes.DataSource_Template);


            var opzioniUtente = getOpzioniUtente(ePPlusHelper, configurazione);

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

        private static UserOptions getOpzioniUtente(EPPlusHelper ePPlusHelper, Configurazione configurazione)
        {
            var worksheetName = Constants.WorksheetNames.DATA_SOURCE_TEMPLATE_CONFIGURATION;
            ThrowExpetionsForMissingWorksheet(ePPlusHelper, worksheetName, FileTypes.DataSource_Template);

            var filtriPossibili = getListaFiltriApplicabili(ePPlusHelper, configurazione);

            return new UserOptions
            {
                Applicablefilters = filtriPossibili
            };
        }

        private static List<FilterItems> getListaFiltriApplicabili(EPPlusHelper ePPlusHelper, Configurazione configurazione)
        {
            var worksheetName = WorksheetNames.DATA_SOURCE_TEMPLATE_CONFIGURATION;

            var filtriPossibili = new List<FilterItems>();

            var rigaCorrente = configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_FILTERS_FIRST_ROW;
            while (true)
            {
                var table = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_FILTERS_TABLE_COL);
                var field = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_FILTERS_FIELD_COL);
                if (string.IsNullOrWhiteSpace(table) || string.IsNullOrWhiteSpace(field))
                { break; }

                filtriPossibili.Add(new FilterItems
                {
                    TableName = table,
                    FieldName = field,
                    Values = new List<string>(),
                    SelectedValues = new List<string>(),
                });

                // passo alla riga successiva
                rigaCorrente++;
            }

            return filtriPossibili;
        }

        #endregion

        private static bool IsBudgetFileOk(string filePath)
        {
            // check file existence

            // check expected worksheet names

            // check expected columns in each worksheet

            // Implement your logic to check if the budget file is OK
            return true;
        }


        static internal EPPlusHelper GetHelperForExistingFile(string filePath, FileTypes fileType)
        {
            var ePPlusHelper = new EPPlusHelper();
            if (!ePPlusHelper.Open(filePath))
            {
                throw new ManagedException(
                    filePath: ePPlusHelper.FilePathInUse,
                    fileType: fileType,
                    //
                    worksheetName: null,
                    cellRow: null,
                    cellColumn: null,
                    valueType: ValueTypes.None,
                    value: null,
                    //
                    errorType: ErrorTypes.UnableToOpenFile,
                    userMessage: UserErrorMessages.UnableToOpenFile
                    );
            }
            return ePPlusHelper;
        }

        static internal void ThrowExpetionsForMissingWorksheet(EPPlusHelper ePPlusHelper, string worksheetName, FileTypes fileType)
        {
            if (!ePPlusHelper.WorksheetExists(worksheetName))
            {
                throw new ManagedException(
                    filePath: ePPlusHelper.FilePathInUse,
                    fileType: fileType,
                    //
                    worksheetName: worksheetName,
                    cellRow: null,
                    cellColumn: null,
                    valueType: ValueTypes.None,
                    value: null,
                    //
                    errorType: ErrorTypes.MissingWorksheet,
                    userMessage: string.Format(UserErrorMessages.MissingWorksheet, worksheetName)
                    );
            }
        }
    }
}
