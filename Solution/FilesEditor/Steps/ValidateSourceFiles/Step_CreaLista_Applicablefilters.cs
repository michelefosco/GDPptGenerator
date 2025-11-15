using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using FilesEditor.Helpers;
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
        public override string StepName => "Step_CreaLista_Applicablefilters";

        internal override void BeforeTask()
        {
            Context.DebugInfoLogger.LogStepContext(StepName, Context);
        }

        internal override void ManageInfoAboutPerformedStepTask(TimeSpan timeSpent)
        {
            Context.DebugInfoLogger.LogPerformance(StepName, timeSpent);
        }

        internal override void AfterTask()
        {
            Context.DebugInfoLogger.LogStepContext(StepName, Context);
        }

        public Step_CreaLista_Applicablefilters(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            CreaLista_Applicablefilters();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }


        private void CreaLista_Applicablefilters()
        {
            //var ePPlusHelper = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(Context.DataSourceFilePath, FileTypes.DataSource);
            var worksheetName = WorksheetNames.DATASOURCE_CONFIGURATION;
            EPPlusHelperUtilities.ThrowExpetionsForMissingWorksheet(Context.DataSourceEPPlusHelper, worksheetName, FileTypes.DataSource);

            // Validazione dei filtri applicabili e lettura dei loro potenziali valori
            Fill_ApplicableFilters_FromConfigurazione(Context.DataSourceEPPlusHelper, Context.Configurazione, Context.ApplicableFilters);
            FillApplicableFiltersWithValues_FromSourceFiles(Context.ApplicableFilters);
        }


        private void Fill_ApplicableFilters_FromConfigurazione(EPPlusHelper ePPlusHelper, Configurazione configurazione, List<InputDataFilters_Item> applicablefilters)
        {
            var worksheetName = WorksheetNames.DATASOURCE_CONFIGURATION;

            var rigaCorrente = configurazione.DATASOURCE_CONFIG_FILTERS_FIRST_DATA_ROW;
            while (true)
            {
                var table = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_CONFIG_FILTERS_TABLE_COL);
                var field = ePPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.DATASOURCE_CONFIG_FILTERS_FIELD_COL);
                if (ValuesHelper.AreAllNulls(table, field))
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
                applicablefilters.Add(new InputDataFilters_Item
                {
                    Table = parsedTable,
                    FieldName = field,
                    PossibleValues = new List<string>(),
                    SelectedValues = new List<string>(),
                });

                // passo alla riga successiva
                rigaCorrente++;
            }
        }

        private void FillApplicableFiltersWithValues_FromSourceFiles(List<InputDataFilters_Item> applicablefilters)
        {


            foreach (var applicablefilter in applicablefilters)
            {
                switch (applicablefilter.Table)
                {
                    case InputDataFilters_Tables.BUDGET:
                        applicablefilter.PossibleValues = GetApplicableFiltersValues_FromSourceFile(
                                ePPlusHelper: Context.BudgetFileEPPlusHelper,
                                // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome
                                worksheetName: null, // WorksheetNames.SOURCEFILE_BUDGET_DATA,
                                fileType: FileTypes.Budget,
                                headersRow: Context.Configurazione.SOURCE_FILES_BUDGET_HEADERS_ROW,
                                headerValue: applicablefilter.FieldName,
                                startCheckHeadersFromColumn: Context.Configurazione.SOURCE_FILES_BUDGET_HEADERS_FIRST_COL
                                );
                        // E' stato necessario rimuovere manualmente il valore "Totale complessivo" per via della struttura insolita dei file Budget e Forecast
                        if (applicablefilter.FieldName == Values.HEADER_BUSINESS)
                        { applicablefilter.PossibleValues.Remove("Totale complessivo"); }

                        break;

                    case InputDataFilters_Tables.FORECAST:
                        applicablefilter.PossibleValues = GetApplicableFiltersValues_FromSourceFile(
                                ePPlusHelper: Context.ForecastFileEPPlusHelper,
                                // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome
                                worksheetName: null, //  WorksheetNames.SOURCEFILE_FORECAST_DATA,
                                fileType: FileTypes.Forecast,
                                headersRow: Context.Configurazione.SOURCE_FILES_FORECAST_HEADERS_ROW,
                                headerValue: applicablefilter.FieldName,
                                startCheckHeadersFromColumn: Context.Configurazione.SOURCE_FILES_FORECAST_HEADERS_FIRST_COL
                                );
                        // E' stato necessario rimuovere manualmente il valore "Totale complessivo" per via della struttura insolita dei file Budget e Forecast
                        if (applicablefilter.FieldName == Values.HEADER_BUSINESS)
                        { applicablefilter.PossibleValues.Remove("Totale complessivo"); }
                        break;

                    case InputDataFilters_Tables.RUNRATE:
                        applicablefilter.PossibleValues = GetApplicableFiltersValues_FromSourceFile(
                                ePPlusHelper: Context.RunRateFileEPPlusHelper,
                                // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome
                                worksheetName: null, //  WorksheetNames.SOURCEFILE_RUN_RATE_DATA,
                                fileType: FileTypes.RunRate,
                                headersRow: Context.Configurazione.SOURCE_FILES_RUNRATE_HEADERS_ROW,
                                headerValue: applicablefilter.FieldName,
                                startCheckHeadersFromColumn: Context.Configurazione.SOURCE_FILES_RUNRATE_HEADERS_FIRST_COL
                                );
                        break;

                    case InputDataFilters_Tables.SUPERDETTAGLI:
                        applicablefilter.PossibleValues = GetApplicableFiltersValues_FromSourceFile(
                                ePPlusHelper: Context.SuperdettagliFileEPPlusHelper,
                                worksheetName: WorksheetNames.SOURCEFILE_SUPERDETTAGLI_DATA,
                                fileType: FileTypes.SuperDettagli,
                                headersRow: Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_ROW,
                                headerValue: applicablefilter.FieldName,
                                startCheckHeadersFromColumn: Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL
                                );
                        break;

                    default:
                        throw new Exception($"Tipo tabella sconosciuto nella configurazione dei filtri: '{applicablefilter.Table}'");
                }
            }
        }

        private List<string> GetApplicableFiltersValues_FromSourceFile(EPPlusHelper ePPlusHelper, string worksheetName, FileTypes fileType, int headersRow, string headerValue, int startCheckHeadersFromColumn)
        {
            //var ePPlusHelper = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(filePath, fileType);

            // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome ad eccezione di Superdettagli
            if (worksheetName == null)
            {
                worksheetName = ePPlusHelper.ExcelPackage.Workbook.Worksheets[1].Name;
            }

            // Controllo che l'header per la colonna che si sta tentando di esista
            var errorMessage = $"The configuration inside the file DataSource includes the filter: '{fileType} - {headerValue}'.\r\nThe worksheet '{worksheetName}' does not contain the corresponding header ('{headerValue}')";
            EPPlusHelperUtilities.ThrowExpetionsForMissingHeader(ePPlusHelper, worksheetName, fileType, headersRow, startCheckHeadersFromColumn, new List<string> { headerValue }, errorMessage);

            var values = ePPlusHelper.GetValuesFromColumnsWithHeader(worksheetName, headersRow, headerValue, true, startCheckHeadersFromColumn, true);

            // Applico gli alias ai cambi "Categoria" e "Business" dei file "Budget" e "Forecast"
            if (fileType == FileTypes.Budget || fileType == FileTypes.Forecast)
            {
                values = values.Select(_ => Context.ApplicaAliasToValue(headerValue, _)).ToList();
            }


            //ePPlusHelper.Close();
            return values.Distinct().OrderBy(n => n).ToList();
        }
    }
}