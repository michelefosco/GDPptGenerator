using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using FilesEditor.Steps.BuildPresentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FilesEditor
{
    public class Editor
    {
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
            var context = new StepContext(configurazione);
            context.SetContextFromInput(buildPresentationInput);
            var stepsSequence = new List<Step_Base>
                {
                    new Step_PredisponiTmpFolder(context),
                    new Step_Start_FileDebugHelper(context),
                    new Step_AggiornaDataSource(context),
                    new Step_CreaFilesImmagini(context),
                    new Step_CreaFilesPowerPoint(context),
                    new Step_EsitoFinaleSuccess(context),
                 };
            return buildPresentationRunStepSequence(stepsSequence, context);
        }

        private static BuildPresentationOutput buildPresentationRunStepSequence(List<Step_Base> stepsSequence, StepContext context)
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
            try
            {
                // Validazioni preliminari sui file di input
                validazioniPreliminari_InputFiles(validaSourceFilesInput, configurazione);

                // Lettura Opzione dal datasource
                var userOptions = getUserOptions(validaSourceFilesInput, configurazione);

                var outout = new ValidaSourceFilesOutput(EsitiFinali.Success)
                {
                    UserOptions = userOptions,
                };
                return outout;
            }
            catch (ManagedException managedException)
            {
                return new ValidaSourceFilesOutput(managedException);
            }
            catch (Exception ex)
            {
                var managedException = new ManagedException(
                    filePath: null,
                    fileType: FileTypes.None,
                    //
                    worksheetName: null,
                    cellRow: null,
                    cellColumn: null,
                    valueHeader: ValueHeaders.None,
                    value: null,
                    //
                    errorType: ErrorTypes.UnhandledException,
                    userMessage: ex.Message + (ex.InnerException != null ? " (" + ex.InnerException.Message + ")" : "")
                    );
                return new ValidaSourceFilesOutput(managedException);
            }
        }
        #endregion


        #region Lettura da DataSource_Template
        private static UserOptions getUserOptions(ValidaSourceFilesInput validaSourceFilesInput, Configurazione configurazione)
        {
            var dataSourceTemplateFile = Path.Combine(validaSourceFilesInput.SourceFilesFolder, FileNames.DATA_SOURCE_TEMPLATE_FILENAME);
            var ePPlusHelper = GetHelperForExistingFile(dataSourceTemplateFile, FileTypes.DataSource_Template);
            var worksheetName = WorksheetNames.DATA_SOURCE_TEMPLATE_CONFIGURATION;
            ThrowExpetionsForMissingWorksheet(ePPlusHelper, worksheetName, FileTypes.DataSource_Template);

            // Validazione dei filtri applicabili e lettura dei loro potenziali valori
            var applicableFilters = getApplicableFilters(ePPlusHelper, configurazione);
            fillApplicableFiltersWithValues(applicableFilters, validaSourceFilesInput, configurazione);

            var slidesToGenerate = getSildeToGenerate(ePPlusHelper, configurazione);

            return new UserOptions
            {
                Applicablefilters = applicableFilters,
                SildeToGenerate = slidesToGenerate
            };
        }


        #region Validazioni preliminari sui file di input
        private static void validazioniPreliminari_InputFiles(ValidaSourceFilesInput validaSourceFilesInput, Configurazione configurazione)
        {
            validazioniPreliminari_InputFiles_SuperDettagli(validaSourceFilesInput.FileSuperDettagliPath, configurazione);
            validazioniPreliminari_InputFiles_Budget(validaSourceFilesInput.FileBudgetPath, configurazione);
            validazioniPreliminari_InputFiles_Forecast(validaSourceFilesInput.FileForecastPath, configurazione);
            validazioniPreliminari_InputFiles_RunRate(validaSourceFilesInput.FileRunRatePath, configurazione);
        }
        private static void validazioniPreliminari_InputFiles_SuperDettagli(string filePath, Configurazione configurazione)
        {
            var fileType = FileTypes.SuperDettagli;
            var worksheetName = WorksheetNames.SUPERDETTAGLI_DATA;
            var headersRow = configurazione.SUPERDETTAGLI_HEADERS_ROW;
            //todo: aggiungere un certo numero di colonne uniche di questo foglio
            var expectedHeadersColumns = new List<string> { "Distinzione produttive indirette vs improduttive" };

            validazioniPreliminari_Comuni(filePath, fileType, worksheetName, headersRow, expectedHeadersColumns);
        }
        private static void validazioniPreliminari_InputFiles_Budget(string filePath, Configurazione configurazione)
        {
            var fileType = FileTypes.Budget;
            var worksheetName = WorksheetNames.BUDGET_DATA;
            var headersRow = configurazione.BUDGET_HEADERS_ROW;
            //todo: aggiungere un certo numero di colonne uniche di questo foglio
            var expectedHeadersColumns = new List<string> { "Macro categoria_new", "Strut.Eng" };

            validazioniPreliminari_Comuni(filePath, fileType, worksheetName, headersRow, expectedHeadersColumns);
        }
        private static void validazioniPreliminari_InputFiles_Forecast(string filePath, Configurazione configurazione)
        {
            var fileType = FileTypes.Forecast;
            var worksheetName = WorksheetNames.FORECAST_DATA;
            var headersRow = configurazione.FORECAST_HEADERS_ROW;
            //todo: aggiungere un certo numero di colonne uniche di questo foglio
            var expectedHeadersColumns = new List<string> { "ENG TOB Totale_", "ENG LPP - INT_" };

            validazioniPreliminari_Comuni(filePath, fileType, worksheetName, headersRow, expectedHeadersColumns);
        }
        private static void validazioniPreliminari_InputFiles_RunRate(string filePath, Configurazione configurazione)
        {
            var fileType = FileTypes.RunRate;
            var worksheetName = WorksheetNames.RUN_RATE_DATA;
            var headersRow = configurazione.RUNRATE_HEADERS_ROW;
            //todo: aggiungere un certo numero di colonne uniche di questo foglio
            var expectedHeadersColumns = new List<string> { "01", "02" };

            validazioniPreliminari_Comuni(filePath, fileType, worksheetName, headersRow, expectedHeadersColumns);
        }
        private static void validazioniPreliminari_Comuni(string filePath, FileTypes fileType, string worksheetName, int headersRow, List<string> expectedHeadersColumns)
        {
            var ePPlusHelper = GetHelperForExistingFile(filePath, fileType);
            // Controllo che ci sia il foglio da cui leggere i dati
            ThrowExpetionsForMissingWorksheet(ePPlusHelper, worksheetName, fileType);

            // Controllo che gli headers corrispondano (almeno in parte a quelli previsti)      
            ThrowExpetionsForMissingHeader(ePPlusHelper, worksheetName, fileType, headersRow, expectedHeadersColumns);
        }
        #endregion


        #region Lettuara dei filtri applicabili     
        private static List<InputDataFilters_Items> getApplicableFilters(EPPlusHelper ePPlusHelper, Configurazione configurazione)
        {
            var worksheetName = WorksheetNames.DATA_SOURCE_TEMPLATE_CONFIGURATION;
            var filtriPossibili = new List<InputDataFilters_Items>();

            var rigaCorrente = configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_FILTERS_FIRST_DATA_ROW;
            while (true)
            {
                var table = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_FILTERS_TABLE_COL);
                var field = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_FILTERS_FIELD_COL);
                if (allNulls(table, field))
                { break; }

                if (!Enum.TryParse(table, out InputDataFilters_Tables parsedTable))
                {
                    //todo: sollevare eccezione Managed
                    throw new Exception("Tipo tabella sconosciuto nella configurazione dei filtri");
                }
                filtriPossibili.Add(new InputDataFilters_Items
                {
                    Table = parsedTable,
                    FieldName = field,
                    Values = new List<string>(),
                    SelectedValues = new List<string>(),
                });

                // passo alla riga successiva
                rigaCorrente++;
            }

            return filtriPossibili;
        }

        private static void fillApplicableFiltersWithValues(List<InputDataFilters_Items> applicablefilters, ValidaSourceFilesInput validaSourceFilesInput, Configurazione configurazione)
        {
            foreach (var applicablefilter in applicablefilters)
            {
                switch (applicablefilter.Table)
                {

                    case InputDataFilters_Tables.SUPERDETTAGLI:
                        applicablefilter.Values = fillApplicableFiltersWithValues_FromFile(
                                filePath: validaSourceFilesInput.FileSuperDettagliPath,
                                worksheetName: WorksheetNames.SUPERDETTAGLI_DATA,
                                fileType: FileTypes.SuperDettagli,
                                headersRow: configurazione.SUPERDETTAGLI_HEADERS_ROW,
                                headerValue: applicablefilter.FieldName);
                        break;
                    case InputDataFilters_Tables.FORECAST:
                        applicablefilter.Values = fillApplicableFiltersWithValues_FromFile(
                                filePath: validaSourceFilesInput.FileForecastPath,
                                worksheetName: WorksheetNames.FORECAST_DATA,
                                fileType: FileTypes.Forecast,
                                headersRow: configurazione.FORECAST_HEADERS_ROW,
                                headerValue: applicablefilter.FieldName);
                        break;
                    case InputDataFilters_Tables.BUDGET:
                        applicablefilter.Values = fillApplicableFiltersWithValues_FromFile(
                                filePath: validaSourceFilesInput.FileBudgetPath,
                                worksheetName: WorksheetNames.BUDGET_DATA,
                                fileType: FileTypes.Budget,
                                headersRow: configurazione.BUDGET_HEADERS_ROW,
                                headerValue: applicablefilter.FieldName);
                        break;
                    case InputDataFilters_Tables.RUNRATE:
                        applicablefilter.Values = fillApplicableFiltersWithValues_FromFile(
                                filePath: validaSourceFilesInput.FileRunRatePath,
                                worksheetName: WorksheetNames.RUN_RATE_DATA,
                                fileType: FileTypes.RunRate,
                                headersRow: configurazione.RUNRATE_HEADERS_ROW,
                                headerValue: applicablefilter.FieldName);
                        break;

                    default:
                        throw new Exception($"Tipo tabella sconosciuto nella configurazione dei filtri: '{applicablefilter.Table}'");

                }
            }
        }

        private static List<string> fillApplicableFiltersWithValues_FromFile(string filePath, string worksheetName, FileTypes fileType, int headersRow, string headerValue)
        {
            var ePPlusHelper = GetHelperForExistingFile(filePath, fileType);
            
            // Controllo che l'header per la colonna che si sta tentando di esista
            ThrowExpetionsForMissingHeader(ePPlusHelper, worksheetName, fileType, headersRow, new List<string> { headerValue });

            var values = ePPlusHelper.GetValuesFromColumnsWithHeader(worksheetName, headersRow, headerValue);
            return values.Distinct().OrderBy(n => n).ToList(); ;
        }

        #endregion

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

                // check sul campo "Content 1"
                ManagedException.ThrowIfMissingMandatoryValue(content1, ePPlusHelper.FilePathInUse, FileTypes.DataSource_Template, worksheetName, rigaCorrente,
                    configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_CONTENT_1_COL,
                    ValueHeaders.SlideTitle);

                var contents = new List<string>() { content1 };
                if (!string.IsNullOrWhiteSpace(content2)) { contents.Add(content2); }
                if (!string.IsNullOrWhiteSpace(content3)) { contents.Add(content3); }
                // Verifico che i valori usati in contents siano validi (ovvero esistano foglio con quel nome)
                foreach (var item in contents)
                {
                    if (!printableWorksheets.Any(n => n.Equals(item, StringComparison.InvariantCultureIgnoreCase)))
                    {
                        //todo: Sollevare eccezione Managed
                        throw new Exception("Elemento da stampare non valido");
                    }
                }




                // con più di un contenuto, il layout diventa obbligatorio
                if (contents.Count > 1)
                {
                    // check sul campo "Layout"
                    ManagedException.ThrowIfMissingMandatoryValue(layout, ePPlusHelper.FilePathInUse, FileTypes.DataSource_Template, worksheetName, rigaCorrente,
                        configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_LAYOUT_COL,
                        ValueHeaders.SlideLayout);
                }
                //todo: chiedere info a Francesco su questo uso del default
                if (layout == null)
                { layout = LayoutTypes.Horizontal.ToString(); }


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
