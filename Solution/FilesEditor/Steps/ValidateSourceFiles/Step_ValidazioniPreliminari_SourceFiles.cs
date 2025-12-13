using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using System;
using System.Collections.Generic;

namespace FilesEditor.Steps.ValidateSourceFiles
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_ValidazioniPreliminari_SourceFiles : StepBase
    {
        internal override string StepName => "Step_ValidazioniPreliminari_SourceFiles";

        public Step_ValidazioniPreliminari_SourceFiles(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            ValidazioniPreliminari_SourceFiles();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        internal void ValidazioniPreliminari_SourceFiles()
        {
            ValidazioniPreliminari_Comuni(
                 datasourceWorksheetName: WorksheetNames.DATASOURCE_BUDGET_DATA,
                 datasourceWorksheetHeadersRow: Context.Configurazione.DATASOURCE_BUDGET_HEADERS_ROW,
                 datasourceWorksheetHeadersFirstColumn: Context.Configurazione.DATASOURCE_BUDGET_HEADERS_FIRST_COL,
                 //
                 //sourceFilePath: Context.FileBudgetPath,
                 sourceFileEPPlusHelper: Context.BudgetFileEPPlusHelper,
                 sourceFileType: FileTypes.Budget,
                 // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome
                 sourceFileWorksheetName: null, // WorksheetNames.SOURCEFILE_BUDGET_DATA,
                 sourceFileHeadersRow: Context.Configurazione.SOURCE_FILES_BUDGET_HEADERS_ROW,
                 sourceFileHeadersFirstColumn: Context.Configurazione.SOURCE_FILES_BUDGET_HEADERS_FIRST_COL,
                 ovverideExpectedHeadersColumns: new List<string>() { "Business", "Categoria" }
                );

            ValidazioniPreliminari_Comuni(
                 datasourceWorksheetName: WorksheetNames.DATASOURCE_FORECAST_DATA,
                 datasourceWorksheetHeadersRow: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_ROW,
                 datasourceWorksheetHeadersFirstColumn: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_FIRST_COL,
                 //
                 //sourceFilePath: Context.FileForecastPath,
                 sourceFileEPPlusHelper: Context.ForecastFileEPPlusHelper,
                 sourceFileType: FileTypes.Forecast,
                 // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome
                 sourceFileWorksheetName: null, // WorksheetNames.SOURCEFILE_FORECAST_DATA,
                 sourceFileHeadersRow: Context.Configurazione.SOURCE_FILES_FORECAST_HEADERS_ROW,
                 sourceFileHeadersFirstColumn: Context.Configurazione.SOURCE_FILES_FORECAST_HEADERS_FIRST_COL,
                 ovverideExpectedHeadersColumns: new List<string>() { "Business", "Categoria" }
                );

            ValidazioniPreliminari_Comuni(
                 datasourceWorksheetName: WorksheetNames.DATASOURCE_RUN_RATE_DATA,
                 datasourceWorksheetHeadersRow: Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_ROW,
                 datasourceWorksheetHeadersFirstColumn: Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_FIRST_COL,
                 //
                 //sourceFilePath: Context.FileRunRatePath,
                 sourceFileEPPlusHelper: Context.RunRateFileEPPlusHelper,
                 sourceFileType: FileTypes.RunRate,
                 // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome
                 sourceFileWorksheetName: null, //  WorksheetNames.SOURCEFILE_RUN_RATE_DATA,
                 sourceFileHeadersRow: Context.Configurazione.SOURCE_FILES_RUNRATE_HEADERS_ROW,
                 sourceFileHeadersFirstColumn: Context.Configurazione.SOURCE_FILES_RUNRATE_HEADERS_FIRST_COL
                );

            ValidazioniPreliminari_Comuni(
                     datasourceWorksheetName: WorksheetNames.DATASOURCE_CN43N_DATA,
                     datasourceWorksheetHeadersRow: Context.Configurazione.DATASOURCE_CN43N_HEADERS_ROW,
                     datasourceWorksheetHeadersFirstColumn: Context.Configurazione.DATASOURCE_CN43N_HEADERS_FIRST_COL,
                     //
                     //sourceFilePath: Context.FileCN43NPath,
                     sourceFileEPPlusHelper: Context.CN43NFileEPPlusHelper,
                     sourceFileType: FileTypes.CN43N,
                     // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome
                     sourceFileWorksheetName: null, //  WorksheetNames.SOURCEFILE_RUN_RATE_DATA,
                     sourceFileHeadersRow: Context.Configurazione.SOURCE_FILES_CN43N_HEADERS_ROW,
                     sourceFileHeadersFirstColumn: Context.Configurazione.SOURCE_FILES_CN43N_HEADERS_FIRST_COL
                    );

            //ValidazioniPreliminari_Comuni(
            //     datasourceWorksheetName: WorksheetNames.DATASOURCE_SUPERDETTAGLI_DATA,
            //     datasourceWorksheetHeadersRow: Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_ROW,
            //     datasourceWorksheetHeadersFirstColumn: Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL,
            //     //
            //     sourceFileEPPlusHelper: Context.SuperdettagliFileEPPlusHelper,
            //     sourceFileType: FileTypes.SuperDettagli,
            //     sourceFileWorksheetName: WorksheetNames.SOURCEFILE_SUPERDETTAGLI_DATA,
            //     sourceFileHeadersRow: Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_ROW,
            //     sourceFileHeadersFirstColumn: Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL
            //    );
        }

        private void ValidazioniPreliminari_Comuni(
            string datasourceWorksheetName,
            int datasourceWorksheetHeadersRow,
            int datasourceWorksheetHeadersFirstColumn,
            //
            //string sourceFilePath,
            EPPlusHelper sourceFileEPPlusHelper,
            FileTypes sourceFileType,
            string sourceFileWorksheetName,
            int sourceFileHeadersRow,
            int sourceFileHeadersFirstColumn,
            List<string> ovverideExpectedHeadersColumns = null
            )
        {

            #region Leggo la lista degli headers richiesti per il dataSource (ovvero le intestazione delle colonne da leggere dai file di input)
            var expectedHeadersColumns = ovverideExpectedHeadersColumns
                                      ?? Context.DataSourceEPPlusHelper.GetHeadersFromRow(datasourceWorksheetName, datasourceWorksheetHeadersRow, datasourceWorksheetHeadersFirstColumn, true);
            #endregion

            #region Verifico che il foglio di input abbia il foglio con tutti gli headers richiesti
            //var sourceFileEPPlusHelper = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(sourceFilePath, sourceFileType);
            // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome
            if (string.IsNullOrWhiteSpace(sourceFileWorksheetName))
            { sourceFileWorksheetName = sourceFileEPPlusHelper.ExcelPackage.Workbook.Worksheets[1].Name; }


            // Controllo che ci sia il foglio da cui leggere i dati
            EPPlusHelperUtilities.ThrowExpetionsForMissingWorksheet(sourceFileEPPlusHelper, sourceFileWorksheetName, sourceFileType);


            // Controllo che gli headers corrispondano (almeno in parte a quelli previsti)      
            EPPlusHelperUtilities.ThrowExpetionsForMissingHeader(sourceFileEPPlusHelper, sourceFileWorksheetName, sourceFileType, sourceFileHeadersRow, sourceFileHeadersFirstColumn, expectedHeadersColumns);
            #endregion
        }
    }
}