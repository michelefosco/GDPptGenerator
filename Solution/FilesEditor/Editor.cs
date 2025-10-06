//using DocumentFormat.OpenXml.Drawing;
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
using System.Linq;

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
            validazioniPreliminari_InputFiles(validaSourceFilesInput, configurazione);


            // Lettura Opzione dal datasource
            var dataSourceTemplateFile = Path.Combine(validaSourceFilesInput.TemplatesFolder, FileNames.DATA_SOURCE_TEMPLATE_FILENAME);
            var ePPlusHelper = GetHelperForExistingFile(dataSourceTemplateFile, FileTypes.DataSource_Template);
            var userOptions = getUserOptions(ePPlusHelper, configurazione);
            ePPlusHelper.Close();





            // lettura info da 1° file

            // lettura info da 2° file

            // lettura info da 3° file

            // lettura info da 4° file

            // verifica applicabilità filtri

            // completa Valori sele

            var outout = new ValidaSourceFilesOutput(EsitiFinali.Success)
            {
                UserOptions = userOptions,
            };
            return outout;
        }
        #endregion


        #region Lettura da DataSource_Template
        private static UserOptions getUserOptions(EPPlusHelper ePPlusHelper, Configurazione configurazione)
        {
            var worksheetName = WorksheetNames.DATA_SOURCE_TEMPLATE_CONFIGURATION;
            ThrowExpetionsForMissingWorksheet(ePPlusHelper, worksheetName, FileTypes.DataSource_Template);

            var filtriPossibili = getListaFiltriApplicabili(ePPlusHelper, configurazione);
            var slidesToGenerate = getSildeToGenerate(ePPlusHelper, configurazione);

            return new UserOptions
            {
                Applicablefilters = filtriPossibili,
                SildeToGenerate = slidesToGenerate
            };
        }

        private static List<FilterItems> getListaFiltriApplicabili(EPPlusHelper ePPlusHelper, Configurazione configurazione)
        {
            var worksheetName = WorksheetNames.DATA_SOURCE_TEMPLATE_CONFIGURATION;

            var filtriPossibili = new List<FilterItems>();

            var rigaCorrente = configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_FILTERS_FIRST_DATA_ROW;
            while (true)
            {
                var table = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_FILTERS_TABLE_COL);
                var field = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_FILTERS_FIELD_COL);
                if (allNulls(table, field))
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

        private static List<SlideToGenerate> getSildeToGenerate(EPPlusHelper ePPlusHelper, Configurazione configurazione)
        {
            var worksheetName = WorksheetNames.DATA_SOURCE_TEMPLATE_CONFIGURATION;


            var printableWorksheets = ePPlusHelper.GetWorksheetNames().Where(n => n.StartsWith(WorksheetNames.DATA_SOURCE_TEMPLATE_PRINTABLE_WORKSHEET_NAME_PREFIX, StringComparison.InvariantCultureIgnoreCase)).ToList();

            //todo: lettura dei workshitt names con elementi stampabili

            var slideToGenerateList = new List<SlideToGenerate>();

            var rigaCorrente = configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_FIRST_DATA_ROW;
            while (true)
            {
                var outputFileName = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_POWERPOINTFILE_COL);
                var title = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_TITLE_COL);
                var content1 = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_CONTENT_1_COL);
                var content2 = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_CONTENT_2_COL);
                var content3 = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_CONTENT_3_COL);
                var layout = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_LAYOUT_COL);

                // mi fermo quando la riga è competamente a null
                if (allNulls(outputFileName, title, content1, content2, content3, layout))
                { break; }

                //verifico i campi obbligatori
                // check sul campo "Powerpoint File"
                ManagedException.ThrowIfMissingMandatoryValue(outputFileName, ePPlusHelper.FilePathInUse, FileTypes.DataSource_Template, worksheetName, rigaCorrente,
                    configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_POWERPOINTFILE_COL,
                    ValueHeaders.TableName);
                // check sul campo "Title"
                ManagedException.ThrowIfMissingMandatoryValue(title, ePPlusHelper.FilePathInUse, FileTypes.DataSource_Template, worksheetName, rigaCorrente,
                    configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_TITLE_COL,
                    ValueHeaders.SlideTitle);
                // check sul campo "Conten 1"
                ManagedException.ThrowIfMissingMandatoryValue(content1, ePPlusHelper.FilePathInUse, FileTypes.DataSource_Template, worksheetName, rigaCorrente,
                    configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_CONTENT_1_COL,
                    ValueHeaders.SlideTitle);

                //todo: chiedere info a Francesco su questo uso del default
                if (layout == null) { layout = LayoutTypes.Horizontal.ToString(); }
                // check sul campo "Layout"
                ManagedException.ThrowIfMissingMandatoryValue(layout, ePPlusHelper.FilePathInUse, FileTypes.DataSource_Template, worksheetName, rigaCorrente,
                    configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_LAYOUT_COL,
                    ValueHeaders.SlideLayout);

                if (Enum.TryParse(layout, out LayoutTypes layoutType))
                {
                    // ok
                }
                else
                {
                    // Sollevare eccezione Managed
                    //todo:
                    throw new Exception("Tipo layout sconosciuto");
                }

                var contents = new List<string>() { content1 };
                if (!string.IsNullOrWhiteSpace(content2)) { contents.Add(content2); }
                if (!string.IsNullOrWhiteSpace(content3)) { contents.Add(content3); }

                // Verifico che i valori usati in contents siano validi (ovvero esistano foglio con quel nome)
                foreach (var item in contents)
                {
                    if (!printableWorksheets.Any(n => n.Equals(item, StringComparison.InvariantCultureIgnoreCase)))
                    {
                        // Sollevare eccezione Managed
                        //todo:
                        throw new Exception("Elemento da stampare non valido");
                    }
                }

                // aggiungo la slide alla lista di quelle lette
                slideToGenerateList.Add(new SlideToGenerate
                {
                    OutputFileName = outputFileName,
                    Title = title,
                    LayoutType = layoutType,
                    Contents = contents
                });

                // passo alla riga successiva
                rigaCorrente++;
            }

            return slideToGenerateList;
        }
        #endregion

        private static void validazioniPreliminari_InputFiles(ValidaSourceFilesInput validaSourceFilesInput, Configurazione configurazione)
        {
            validazioniPreliminari_InputFiles_SuperDettagli(validaSourceFilesInput.FileSuperDettagliPath, configurazione);
        }

        private static void validazioniPreliminari_InputFiles_SuperDettagli(string filePath, Configurazione configurazione)
        {
            var fileType = FileTypes.SuperDettagli;
            var ePPlusHelper = GetHelperForExistingFile(filePath, fileType);
            //
            var worksheetName = WorksheetNames.SUPERDETTAGLI_DATA;

            // Controllo che ci sia il foglio da cui leggere i dati
            ThrowExpetionsForMissingWorksheet(ePPlusHelper, worksheetName, fileType);

            // Controllo che gli headers corrispondano (almeno in parte a quelli previsti)
            //todo: aggiungere un certo numero di colonne uniche di questo foglio
            var expectedColumns = new List<string> { "Distinzione produttive indirette vs improduttive"};
            ThrowExpetionsForMissingHeader(ePPlusHelper, worksheetName, fileType, configurazione.SUPERDETTAGLI_HEADERS_ROW, expectedColumns);
        }


    


        #region Utilities
        private static bool allNulls(object obj1, object obj2, object obj3 = null, object obj4 = null, object obj5 = null, object obj6 = null)
        {
            return (obj1 == null && obj2 == null && obj3 == null && obj4 == null && obj5 == null && obj6 == null);

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
                    valueHeader: ValueHeaders.None,
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
                    valueHeader: ValueHeaders.None,
                    value: null,
                    //
                    errorType: ErrorTypes.MissingWorksheet,
                    userMessage: string.Format(UserErrorMessages.MissingWorksheet, worksheetName)
                    );
            }
        }
        private static void ThrowExpetionsForMissingHeader(EPPlusHelper ePPlusHelper, string worksheetName, FileTypes fileType, int rowWithHeaders, List<string> expectedColumns)
        {
            var columnsList = ePPlusHelper.GetHeaders(worksheetName, rowWithHeaders);
            foreach (var expectedColumn in expectedColumns)
            {
                if (!columnsList.Any(_ => _.Equals(expectedColumn, StringComparison.InvariantCultureIgnoreCase)))
                {
                    throw new ManagedException(
                            filePath: ePPlusHelper.FilePathInUse,
                            fileType: fileType,
                            //
                            worksheetName: worksheetName,
                            cellRow: rowWithHeaders,
                            cellColumn: null,
                            valueHeader: ValueHeaders.None,
                            value: null,
                            //
                            errorType: ErrorTypes.MissingValue,
                            userMessage: $"The file '{fileType}' does not have one of the expected headers ('{expectedColumn}')"
                            );
                }
            }
        }
        #endregion
    }
}
