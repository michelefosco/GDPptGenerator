using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;

namespace FilesEditor.Steps.ValidateSourceFiles
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_LetturaUserOptions : ValidateSourceFiles_StepBase
    {
        public Step_LetturaUserOptions(StepContext context) : base(context)
        { }

        internal override ValidaSourceFilesOutput DoSpecificTask()
        {
            getUserOptionsFromDataSource();
            return null; // Step intermedio, non ritorna alcun esito
        }


        private  void getUserOptionsFromDataSource()
        {
            var dataSourceTemplateFile = Path.Combine(Context.SourceFilesFolder, FileNames.DATA_SOURCE_TEMPLATE_FILENAME);
            var ePPlusHelper = GetHelperForExistingFile(dataSourceTemplateFile, FileTypes.DataSource_Template);
            var worksheetName = WorksheetNames.DATA_SOURCE_TEMPLATE_CONFIGURATION;
            ThrowExpetionsForMissingWorksheet(ePPlusHelper, worksheetName, FileTypes.DataSource_Template);

            // Validazione dei filtri applicabili e lettura dei loro potenziali valori
            var applicableFilters = getApplicableFilters(ePPlusHelper, Context.Configurazione);
            fillApplicableFiltersWithValues(applicableFilters);

            var slidesToGenerate = getSildeToGenerate(ePPlusHelper, Context.Configurazione);

            Context.UserOptions = new UserOptions
            {
                Applicablefilters = applicableFilters,
                SildeToGenerate = slidesToGenerate
            };
        }



        #region Lettuara dei filtri applicabili     
        private  List<InputDataFilters_Items> getApplicableFilters(EPPlusHelper ePPlusHelper, Configurazione configurazione)
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

        private  void fillApplicableFiltersWithValues(List<InputDataFilters_Items> applicablefilters)
        {
            foreach (var applicablefilter in applicablefilters)
            {
                switch (applicablefilter.Table)
                {

                    case InputDataFilters_Tables.SUPERDETTAGLI:
                        applicablefilter.Values = fillApplicableFiltersWithValues_FromFile(
                                filePath: Context.FileSuperDettagliPath,
                                worksheetName: WorksheetNames.SUPERDETTAGLI_DATA,
                                fileType: FileTypes.SuperDettagli,
                                headersRow: Context.Configurazione.SUPERDETTAGLI_HEADERS_ROW,
                                headerValue: applicablefilter.FieldName);
                        break;
                    case InputDataFilters_Tables.FORECAST:
                        applicablefilter.Values = fillApplicableFiltersWithValues_FromFile(
                                filePath: Context.FileForecastPath,
                                worksheetName: WorksheetNames.FORECAST_DATA,
                                fileType: FileTypes.Forecast,
                                headersRow: Context.Configurazione.FORECAST_HEADERS_ROW,
                                headerValue: applicablefilter.FieldName);
                        break;
                    case InputDataFilters_Tables.BUDGET:
                        applicablefilter.Values = fillApplicableFiltersWithValues_FromFile(
                                filePath: Context.FileBudgetPath,
                                worksheetName: WorksheetNames.BUDGET_DATA,
                                fileType: FileTypes.Budget,
                                headersRow: Context.Configurazione.BUDGET_HEADERS_ROW,
                                headerValue: applicablefilter.FieldName);
                        break;
                    case InputDataFilters_Tables.RUNRATE:
                        applicablefilter.Values = fillApplicableFiltersWithValues_FromFile(
                                filePath: Context.FileRunRatePath,
                                worksheetName: WorksheetNames.RUN_RATE_DATA,
                                fileType: FileTypes.RunRate,
                                headersRow: Context.Configurazione.RUNRATE_HEADERS_ROW,
                                headerValue: applicablefilter.FieldName);
                        break;

                    default:
                        throw new Exception($"Tipo tabella sconosciuto nella configurazione dei filtri: '{applicablefilter.Table}'");

                }
            }
        }

        private  List<string> fillApplicableFiltersWithValues_FromFile(string filePath, string worksheetName, FileTypes fileType, int headersRow, string headerValue)
        {
            var ePPlusHelper = GetHelperForExistingFile(filePath, fileType);

            // Controllo che l'header per la colonna che si sta tentando di esista
            ThrowExpetionsForMissingHeader(ePPlusHelper, worksheetName, fileType, headersRow, new List<string> { headerValue });

            var values = ePPlusHelper.GetValuesFromColumnsWithHeader(worksheetName, headersRow, headerValue);
            return values.Distinct().OrderBy(n => n).ToList(); ;
        }

        #endregion

        private  List<SlideToGenerate> getSildeToGenerate(EPPlusHelper ePPlusHelper, Configurazione configurazione)
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
    }
}