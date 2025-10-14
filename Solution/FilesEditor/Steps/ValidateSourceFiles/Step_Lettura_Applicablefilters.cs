using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Enums;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FilesEditor.Steps.ValidateSourceFiles
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_Lettura_Applicablefilters : StepBase
    {
        public Step_Lettura_Applicablefilters(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            getDataFromDataSourceFile();
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }


        private  void getDataFromDataSourceFile()
        {
            var dataSourceTemplateFile = Path.Combine(Context.SourceFilesFolder, FileNames.DATA_SOURCE_TEMPLATE_FILENAME);
            var ePPlusHelper = GetHelperForExistingFile(dataSourceTemplateFile, FileTypes.DataSource_Template);
            var worksheetName = WorksheetNames.DATA_SOURCE_TEMPLATE_CONFIGURATION;
            ThrowExpetionsForMissingWorksheet(ePPlusHelper, worksheetName, FileTypes.DataSource_Template);

            // Validazione dei filtri applicabili e lettura dei loro potenziali valori
            var applicableFilters = getApplicableFilters(ePPlusHelper, Context.Configurazione);
            fillApplicableFiltersWithValues(applicableFilters);

            Context.Applicablefilters = applicableFilters;
        }


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
    }
}