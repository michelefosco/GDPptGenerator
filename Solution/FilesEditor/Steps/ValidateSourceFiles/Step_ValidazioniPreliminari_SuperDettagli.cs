using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using System;


namespace FilesEditor.Steps.ValidateSourceFiles
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_ValidazioniPreliminari_SuperDettagli : StepBase
    {
        public override string StepName => "Step_ValidazioniPreliminari_SuperDettagli";

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

        public Step_ValidazioniPreliminari_SuperDettagli(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            ValidazioniPreliminari_SuperDettagli();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        internal void ValidazioniPreliminari_SuperDettagli()
        {
            var sourceFilePath = Context.FileSuperDettagliPath;
            var sourceWorksheetName = WorksheetNames.SOURCEFILE_SUPERDETTAGLI_DATA;
            var souceHeadersRow = Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_ROW;
            var sourceHeadersFirstColumn = Context.Configurazione.SOURCE_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL;
            //
            var destWorksheetName = WorksheetNames.DATASOURCE_SUPERDETTAGLI_DATA;
            var destHeadersRow = Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_ROW;
            var destHeadersFirstColumn = Context.Configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL;

            // Controllo che ci sia il foglio da cui leggere i dati
            EPPlusHelperUtilities.ThrowExpetionsForMissingWorksheet(Context.SuperdettagliFileEPPlusHelper, sourceWorksheetName, FileTypes.SuperDettagli);

            #region Lettura degli headers del foglio sorgente e foglio destinazione
            var sourceHeaders = Context.SuperdettagliFileEPPlusHelper.GetHeadersFromRow(sourceWorksheetName, souceHeadersRow, sourceHeadersFirstColumn, true);
            var destHeaders = Context.DataSourceEPPlusHelper.GetHeadersFromRow(destWorksheetName, destHeadersRow, destHeadersFirstColumn, true);
            #endregion

            #region Verifico che tutti gli headers necessari per la destinazione siano presenti nella sorgente (e nella giusta posizione)
            if (sourceHeaders.Count < destHeaders.Count)
            {
                throw new ManagedException(
                        filePath: sourceFilePath,
                        fileType: FileTypes.SuperDettagli,
                        //
                        worksheetName: sourceWorksheetName,
                        cellRow: null,
                        cellColumn: null,
                        valueHeader: ValueHeaders.None,
                        value: null,
                        //
                        errorType: ErrorTypes.MissingValue,
                        userMessage: "There are fewer headers in the Superdettagli file than required to complete the DataSource file.\nAll headers in the DataSource file(worksheet Superdettagli) must also be present in the Superdettagli file, in the same order."
                        );
            }

            for (var j = 0; j < destHeaders.Count; j++)
            {
                if (!sourceHeaders[j].Equals(destHeaders[j], StringComparison.InvariantCultureIgnoreCase))
                {
                    throw new ManagedException(
                        filePath: sourceFilePath,
                        fileType: FileTypes.SuperDettagli,
                        //
                        worksheetName: sourceWorksheetName,
                        cellRow: souceHeadersRow,
                        cellColumn: null,
                        valueHeader: ValueHeaders.None,
                        value: destHeaders[j],
                        //
                        errorType: ErrorTypes.MissingValue,
                        userMessage: $"The header '{destHeaders[j]}' required for the file Datasource is missing (or located in the wrong position) in the file 'Superdettagli'.\nAll headers in the DataSource file (worksheet Superdettagli) must also be present in the Superdettagli file, in the same order."
                        );
                }
            }
            #endregion
        }
    }
}