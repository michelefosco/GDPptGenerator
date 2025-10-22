using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Enums;
using FilesEditor.Helpers;

namespace FilesEditor.Steps.ValidateSourceFiles
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_ValidazioniPreliminari_InputFiles : StepBase
    {
        public Step_ValidazioniPreliminari_InputFiles(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            Context.DebugInfoLogger.LogStepContext("Step_ValidazioniPreliminari_InputFiles", Context);
            validazioniPreliminari_InputFiles();
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        internal void validazioniPreliminari_InputFiles()
        {
            validazioniPreliminari_Comuni(
                 datasourceWorksheetName: WorksheetNames.DATASOURCE_BUDGET_DATA,
                 datasourceWorksheetHeadersRow: Context.Configurazione.DATASOURCE_BUDGET_HEADERS_ROW,
                 datasourceWorksheetHeadersFirstColumn: Context.Configurazione.DATASOURCE_BUDGET_HEADERS_FIRST_COL,
                 //
                 inputFilePath: Context.FileBudgetPath,
                 inputFileType: FileTypes.Budget,
                 inputFileWorksheetName: WorksheetNames.INPUTFILES_BUDGET_DATA,
                 inputFileHeadersRow: Context.Configurazione.INPUT_FILES_BUDGET_HEADERS_ROW
                );

            validazioniPreliminari_Comuni(
                 datasourceWorksheetName: WorksheetNames.DATASOURCE_FORECAST_DATA,
                 datasourceWorksheetHeadersRow: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_ROW,
                 datasourceWorksheetHeadersFirstColumn: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_FIRST_COL,
                 //
                 inputFilePath: Context.FileForecastPath,
                 inputFileType: FileTypes.Forecast,
                 inputFileWorksheetName: WorksheetNames.INPUTFILES_FORECAST_DATA,
                 inputFileHeadersRow: Context.Configurazione.INPUT_FILES_FORECAST_HEADERS_ROW
                );

            validazioniPreliminari_Comuni(
                 datasourceWorksheetName: WorksheetNames.DATASOURCE_RUN_RATE_DATA,
                 datasourceWorksheetHeadersRow: Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_ROW,
                 datasourceWorksheetHeadersFirstColumn: Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_FIRST_COL,
                 //
                 inputFilePath: Context.FileRunRatePath,
                 inputFileType: FileTypes.RunRate,
                 inputFileWorksheetName: WorksheetNames.INPUTFILES_RUN_RATE_DATA,
                 inputFileHeadersRow: Context.Configurazione.INPUT_FILES_RUNRATE_HEADERS_ROW
                );

            validazioniPreliminari_Comuni(
                 datasourceWorksheetName: WorksheetNames.DATASOURCE_SUPERDETTAGLI_DATA,
                 datasourceWorksheetHeadersRow: Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_ROW,
                 datasourceWorksheetHeadersFirstColumn: Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL,
                 //
                 inputFilePath: Context.FileSuperDettagliPath,
                 inputFileType: FileTypes.SuperDettagli,
                 inputFileWorksheetName: WorksheetNames.INPUTFILES_SUPERDETTAGLI_DATA,
                 inputFileHeadersRow: Context.Configurazione.INPUT_FILES_SUPERDETTAGLI_HEADERS_ROW
                );
        }

        private void validazioniPreliminari_Comuni(
            string datasourceWorksheetName,
            int datasourceWorksheetHeadersRow,
            int datasourceWorksheetHeadersFirstColumn,
            //
            string inputFilePath,
            FileTypes inputFileType,
            string inputFileWorksheetName,
            int inputFileHeadersRow)
        {
            #region Leggo la lista degli headers richiesti per il dataSource (ovvero le intestazione delle colonne da leggere dai file di input)
           // var dataSourceEPPlusHelper = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(Context.DataSourceFilePath, FileTypes.DataSource);
            var expectedHeadersColumns = Context.ePPlusHelperDataSource.GetHeaders(datasourceWorksheetName, datasourceWorksheetHeadersRow, datasourceWorksheetHeadersFirstColumn);
            #endregion

            #region Verifico che il foglio di input abbia il foglio con tutti gli headers richiesti
            var inputFileEPPlusHelper = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(inputFilePath, inputFileType);

            // Controllo che ci sia il foglio da cui leggere i dati
            EPPlusHelperUtilities.ThrowExpetionsForMissingWorksheet(inputFileEPPlusHelper, inputFileWorksheetName, inputFileType);

            // Controllo che gli headers corrispondano (almeno in parte a quelli previsti)      
            EPPlusHelperUtilities.ThrowExpetionsForMissingHeader(inputFileEPPlusHelper, inputFileWorksheetName, inputFileType, inputFileHeadersRow, expectedHeadersColumns);
            #endregion
        }
    }
}