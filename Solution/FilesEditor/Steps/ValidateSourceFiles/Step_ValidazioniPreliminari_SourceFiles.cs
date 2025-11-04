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
        public override string StepName => "Step_ValidazioniPreliminari_SourceFiles";

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

        public Step_ValidazioniPreliminari_SourceFiles(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoStepTask()
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
                 sourceFilePath: Context.FileBudgetPath,
                 sourceFileType: FileTypes.Budget,
                 sourceFileWorksheetName: WorksheetNames.SOURCEFILE_BUDGET_DATA,
                 sourceFileHeadersRow: Context.Configurazione.SOURCE_FILES_BUDGET_HEADERS_ROW,
                 ovverideExpectedHeadersColumns: new List<string>() {"Business", "Categoria" }
                );

            ValidazioniPreliminari_Comuni(
                 datasourceWorksheetName: WorksheetNames.DATASOURCE_FORECAST_DATA,
                 datasourceWorksheetHeadersRow: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_ROW,
                 datasourceWorksheetHeadersFirstColumn: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_FIRST_COL,
                 //
                 sourceFilePath: Context.FileForecastPath,
                 sourceFileType: FileTypes.Forecast,
                 sourceFileWorksheetName: WorksheetNames.SOURCEFILE_FORECAST_DATA,
                 sourceFileHeadersRow: Context.Configurazione.SOURCE_FILES_FORECAST_HEADERS_ROW,
                 ovverideExpectedHeadersColumns: new List<string>() { "Business", "Categoria" }
                );

            ValidazioniPreliminari_Comuni(
                 datasourceWorksheetName: WorksheetNames.DATASOURCE_RUN_RATE_DATA,
                 datasourceWorksheetHeadersRow: Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_ROW,
                 datasourceWorksheetHeadersFirstColumn: Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_FIRST_COL,
                 //
                 sourceFilePath: Context.FileRunRatePath,
                 sourceFileType: FileTypes.RunRate,
                 sourceFileWorksheetName: WorksheetNames.SOURCEFILE_RUN_RATE_DATA,
                 sourceFileHeadersRow: Context.Configurazione.SOURCE_FILES_RUNRATE_HEADERS_ROW
                );

            ValidazioniPreliminari_Comuni(
                 datasourceWorksheetName: WorksheetNames.DATASOURCE_SUPERDETTAGLI_DATA,
                 datasourceWorksheetHeadersRow: Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_ROW,
                 datasourceWorksheetHeadersFirstColumn: Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL,
                 //
                 sourceFilePath: Context.FileSuperDettagliPath,
                 sourceFileType: FileTypes.SuperDettagli,
                 sourceFileWorksheetName: WorksheetNames.SOURCEFILE_SUPERDETTAGLI_DATA,
                 sourceFileHeadersRow: Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_ROW
                );
        }

        private void ValidazioniPreliminari_Comuni(
            string datasourceWorksheetName,
            int datasourceWorksheetHeadersRow,
            int datasourceWorksheetHeadersFirstColumn,
            //
            string sourceFilePath,
            FileTypes sourceFileType,
            string sourceFileWorksheetName,
            int sourceFileHeadersRow,
            List<string> ovverideExpectedHeadersColumns = null
            )
        {
            #region Leggo la lista degli headers richiesti per il dataSource (ovvero le intestazione delle colonne da leggere dai file di input)
            // var dataSourceEPPlusHelper = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(Context.DataSourceFilePath, FileTypes.DataSource);
            var expectedHeadersColumns = (ovverideExpectedHeadersColumns != null)
                ? ovverideExpectedHeadersColumns
                : Context.EpplusHelperDataSource.GetHeaders(datasourceWorksheetName, datasourceWorksheetHeadersRow, datasourceWorksheetHeadersFirstColumn);
            #endregion

            #region Verifico che il foglio di input abbia il foglio con tutti gli headers richiesti
            var sourceFileEPPlusHelper = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(sourceFilePath, sourceFileType);

            // Controllo che ci sia il foglio da cui leggere i dati
            EPPlusHelperUtilities.ThrowExpetionsForMissingWorksheet(sourceFileEPPlusHelper, sourceFileWorksheetName, sourceFileType);

            // Controllo che gli headers corrispondano (almeno in parte a quelli previsti)      
            EPPlusHelperUtilities.ThrowExpetionsForMissingHeader(sourceFileEPPlusHelper, sourceFileWorksheetName, sourceFileType, sourceFileHeadersRow, expectedHeadersColumns);
            #endregion
        }
    }
}