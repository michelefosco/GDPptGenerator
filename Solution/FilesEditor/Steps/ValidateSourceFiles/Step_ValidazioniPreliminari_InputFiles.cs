using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Enums;
using System.Collections.Generic;

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
                 datasourceWorksheetName: WorksheetNames.DATA_SOURCE_BUDGET_DATA,
                 datasourceWorksheetHeadersRow: Context.Configurazione.DATASOURCE_BUDGET_HEADERS_ROW,
                 datasourceWorksheetHeadersFirstColumn: Context.Configurazione.DATASOURCE_BUDGET_HEADERS_FIRST_COL,
                 //
                 inputFilePath: Context.FileBudgetPath,
                 inputFileType: FileTypes.Budget,
                 inputFileWorksheetName: WorksheetNames.BUDGET_DATA,
                 inputFileHeadersRow: Context.Configurazione.INPUT_FILES_BUDGET_HEADERS_ROW
                );

            validazioniPreliminari_Comuni(
                 datasourceWorksheetName: WorksheetNames.DATA_SOURCE_FORECAST_DATA,
                 datasourceWorksheetHeadersRow: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_ROW,
                 datasourceWorksheetHeadersFirstColumn: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_FIRST_COL,
                 //
                 inputFilePath: Context.FileForecastPath,
                 inputFileType: FileTypes.Forecast,
                 inputFileWorksheetName: WorksheetNames.FORECAST_DATA,
                 inputFileHeadersRow: Context.Configurazione.INPUT_FILES_FORECAST_HEADERS_ROW
                );

            validazioniPreliminari_Comuni(
                 datasourceWorksheetName: WorksheetNames.DATA_SOURCE_RUN_RATE_DATA,
                 datasourceWorksheetHeadersRow: Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_ROW,
                 datasourceWorksheetHeadersFirstColumn: Context.Configurazione.DATASOURCE_RUNRATE_HEADERS_FIRST_COL,
                 //
                 inputFilePath: Context.FileRunRatePath,
                 inputFileType: FileTypes.RunRate,
                 inputFileWorksheetName: WorksheetNames.RUN_RATE_DATA,
                 inputFileHeadersRow: Context.Configurazione.INPUT_FILES_RUNRATE_HEADERS_ROW
                );

            validazioniPreliminari_Comuni(
                 datasourceWorksheetName: WorksheetNames.DATA_SOURCE_SUPERDETTAGLI_DATA,
                 datasourceWorksheetHeadersRow: Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_ROW,
                 datasourceWorksheetHeadersFirstColumn: Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL,
                 //
                 inputFilePath: Context.FileSuperDettagliPath,
                 inputFileType: FileTypes.SuperDettagli,
                 inputFileWorksheetName: WorksheetNames.SUPERDETTAGLI_DATA,
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
            // leggo la lista degli headers richiesti (ovvero le intestazione delle colonne da leggere dai file di input)
            var dataSourceEPPlusHelper = GetHelperForExistingFile(Context.DataSourceFilePath, FileTypes.DataSource);
            var expectedHeadersColumns = dataSourceEPPlusHelper.GetHeaders(datasourceWorksheetName, datasourceWorksheetHeadersRow, datasourceWorksheetHeadersFirstColumn);


            var inputFileEPPlusHelper = GetHelperForExistingFile(inputFilePath, inputFileType);

            // Controllo che ci sia il foglio da cui leggere i dati
            ThrowExpetionsForMissingWorksheet(inputFileEPPlusHelper, inputFileWorksheetName, inputFileType);

            // Controllo che gli headers corrispondano (almeno in parte a quelli previsti)      
            ThrowExpetionsForMissingHeader(inputFileEPPlusHelper, inputFileWorksheetName, inputFileType, inputFileHeadersRow, expectedHeadersColumns);
        }
    }
}