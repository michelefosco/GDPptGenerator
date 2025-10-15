using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.MethodsArgs;
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
            validazioniPreliminari_InputFiles_SuperDettagli(Context.FileSuperDettagliPath, Context.Configurazione);
            validazioniPreliminari_InputFiles_Budget(Context.FileBudgetPath, Context.Configurazione);
            validazioniPreliminari_InputFiles_Forecast(Context.FileForecastPath, Context.Configurazione);
            validazioniPreliminari_InputFiles_RunRate(Context.FileRunRatePath, Context.Configurazione);
        }

        private void validazioniPreliminari_InputFiles_SuperDettagli(string filePath, Configurazione configurazione)
        {
            var fileType = FileTypes.SuperDettagli;
            var worksheetName = WorksheetNames.SUPERDETTAGLI_DATA;
            var headersRow = configurazione.INPUT_FILES_SUPERDETTAGLI_HEADERS_ROW;
            //todo: aggiungere un certo numero di colonne uniche di questo foglio
            var expectedHeadersColumns = new List<string> { "Distinzione produttive indirette vs improduttive" };

            validazioniPreliminari_Comuni(filePath, fileType, worksheetName, headersRow, expectedHeadersColumns);
        }
        private void validazioniPreliminari_InputFiles_Budget(string filePath, Configurazione configurazione)
        {
            var fileType = FileTypes.Budget;
            var worksheetName = WorksheetNames.BUDGET_DATA;
            var headersRow = configurazione.INPUT_FILES_BUDGET_HEADERS_ROW;
            //todo: aggiungere un certo numero di colonne uniche di questo foglio
            var expectedHeadersColumns = new List<string> { "Macro categoria_new", "Strut.Eng" };

            validazioniPreliminari_Comuni(filePath, fileType, worksheetName, headersRow, expectedHeadersColumns);
        }
        private void validazioniPreliminari_InputFiles_Forecast(string filePath, Configurazione configurazione)
        {
            var fileType = FileTypes.Forecast;
            var worksheetName = WorksheetNames.FORECAST_DATA;
            var headersRow = configurazione.INPUT_FILES_FORECAST_HEADERS_ROW;
            //todo: aggiungere un certo numero di colonne uniche di questo foglio
            var expectedHeadersColumns = new List<string> { "ENG TOB Totale_", "ENG LPP - INT_" };

            validazioniPreliminari_Comuni(filePath, fileType, worksheetName, headersRow, expectedHeadersColumns);
        }
        private void validazioniPreliminari_InputFiles_RunRate(string filePath, Configurazione configurazione)
        {
            var fileType = FileTypes.RunRate;
            var worksheetName = WorksheetNames.RUN_RATE_DATA;
            var headersRow = configurazione.INPUT_FILES_RUNRATE_HEADERS_ROW;
            //todo: aggiungere un certo numero di colonne uniche di questo foglio
            var expectedHeadersColumns = new List<string> { "01", "02" };

            validazioniPreliminari_Comuni(filePath, fileType, worksheetName, headersRow, expectedHeadersColumns);
        }

        private void validazioniPreliminari_Comuni(string filePath, FileTypes fileType, string worksheetName, int headersRow, List<string> expectedHeadersColumns)
        {
            var ePPlusHelper = GetHelperForExistingFile(filePath, fileType);
            // Controllo che ci sia il foglio da cui leggere i dati
            ThrowExpetionsForMissingWorksheet(ePPlusHelper, worksheetName, fileType);

            // Controllo che gli headers corrispondano (almeno in parte a quelli previsti)      
            ThrowExpetionsForMissingHeader(ePPlusHelper, worksheetName, fileType, headersRow, expectedHeadersColumns);
        }
    }
}