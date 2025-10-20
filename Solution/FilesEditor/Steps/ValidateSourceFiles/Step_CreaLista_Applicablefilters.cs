using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FilesEditor.Steps.ValidateSourceFiles
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_CreaLista_Applicablefilters : StepBase
    {
        public Step_CreaLista_Applicablefilters(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            Context.DebugInfoLogger.LogStepContext("Step_CreaLista_Applicablefilters", Context);
            creaLista_Applicablefilters();
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }


        private void creaLista_Applicablefilters()
        {
            var ePPlusHelper = GetHelperForExistingFile(Context.DataSourceFilePath, FileTypes.DataSource);
            var worksheetName = WorksheetNames.DATA_SOURCE_CONFIGURATION;
            ThrowExpetionsForMissingWorksheet(ePPlusHelper, worksheetName, FileTypes.DataSource);

            // Validazione dei filtri applicabili e lettura dei loro potenziali valori
            var applicableFilters = getApplicableFilters(ePPlusHelper, Context.Configurazione);
            fillApplicableFiltersWithValues(applicableFilters);

            Context.Applicablefilters = applicableFilters;
        }


        private List<InputDataFilters_Item> getApplicableFilters(EPPlusHelper ePPlusHelper, Configurazione configurazione)
        {
            var worksheetName = WorksheetNames.DATA_SOURCE_CONFIGURATION;
            var filtriPossibili = new List<InputDataFilters_Item>();

            var rigaCorrente = configurazione.DATASOURCE_CONFIG_FILTERS_FIRST_DATA_ROW;
            while (true)
            {
                var table = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_CONFIG_FILTERS_TABLE_COL);
                var field = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_CONFIG_FILTERS_FIELD_COL);
                if (allNulls(table, field))
                { break; }

                if (!Enum.TryParse(table, out InputDataFilters_Tables parsedTable))
                {
                    throw new ManagedException(
                        filePath: ePPlusHelper.FilePathInUse,
                        fileType: FileTypes.DataSource,
                        //
                        worksheetName: worksheetName,
                        cellRow: rigaCorrente,
                        cellColumn: configurazione.DATASOURCE_CONFIG_FILTERS_TABLE_COL,
                        valueHeader: ValueHeaders.None,
                        value: table,
                        //
                        errorType: ErrorTypes.InvalidValue,
                        userMessage: string.Format(UserErrorMessages.InvalidValue, table)
                        );
                }
                filtriPossibili.Add(new InputDataFilters_Item
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

        private void fillApplicableFiltersWithValues(List<InputDataFilters_Item> applicablefilters)
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
                                headersRow: Context.Configurazione.INPUT_FILES_SUPERDETTAGLI_HEADERS_ROW,
                                headerValue: applicablefilter.FieldName);
                        break;
                    case InputDataFilters_Tables.FORECAST:
                        applicablefilter.Values = fillApplicableFiltersWithValues_FromFile(
                                filePath: Context.FileForecastPath,
                                worksheetName: WorksheetNames.FORECAST_DATA,
                                fileType: FileTypes.Forecast,
                                headersRow: Context.Configurazione.INPUT_FILES_FORECAST_HEADERS_ROW,
                                headerValue: applicablefilter.FieldName);
                        break;
                    case InputDataFilters_Tables.BUDGET:
                        applicablefilter.Values = fillApplicableFiltersWithValues_FromFile(
                                filePath: Context.FileBudgetPath,
                                worksheetName: WorksheetNames.BUDGET_DATA,
                                fileType: FileTypes.Budget,
                                headersRow: Context.Configurazione.INPUT_FILES_BUDGET_HEADERS_ROW,
                                headerValue: applicablefilter.FieldName);
                        break;
                    case InputDataFilters_Tables.RUNRATE:
                        applicablefilter.Values = fillApplicableFiltersWithValues_FromFile(
                                filePath: Context.FileRunRatePath,
                                worksheetName: WorksheetNames.RUN_RATE_DATA,
                                fileType: FileTypes.RunRate,
                                headersRow: Context.Configurazione.INPUT_FILES_RUNRATE_HEADERS_ROW,
                                headerValue: applicablefilter.FieldName);
                        break;

                    default:
                        throw new Exception($"Tipo tabella sconosciuto nella configurazione dei filtri: '{applicablefilter.Table}'");

                }
            }
        }

        private List<string> fillApplicableFiltersWithValues_FromFile(string filePath, string worksheetName, FileTypes fileType, int headersRow, string headerValue)
        {
            var ePPlusHelper = GetHelperForExistingFile(filePath, fileType);

            // Controllo che l'header per la colonna che si sta tentando di esista
            var errorMessage = $"The configuration inside the file DataSource includes the filter: '{fileType} - {headerValue}'.\r\nThe worksheet '{worksheetName}' does not contain the corresponding header ('{headerValue}')";
            ThrowExpetionsForMissingHeader(ePPlusHelper, worksheetName, fileType, headersRow, new List<string> { headerValue }, errorMessage);

            var values = ePPlusHelper.GetValuesFromColumnsWithHeader(worksheetName, headersRow, headerValue);
            return values.Distinct().OrderBy(n => n).ToList(); ;
        }
    }
}