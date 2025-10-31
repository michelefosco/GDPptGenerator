using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// Avvio il DataSource EPPlusHelper per assicurarmi che il file sia editabile
    /// </summary>
    internal class Step_VerificaEditabilita_DataSource_File : StepBase
    {
        public Step_VerificaEditabilita_DataSource_File(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            tryToSaveDataSourceFile();

            Context.DebugInfoLogger.LogStepContext("Step_VerificaEditabilita_DataSource_File", Context);

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void tryToSaveDataSourceFile()
        {
            var testSaveDataSourcePassed = Context.EpplusHelperDataSource.Save();
            if (!testSaveDataSourcePassed)
            {
                throw new ManagedException(
                    filePath: Context.EpplusHelperDataSource.FilePathInUse,
                    fileType: FileTypes.DataSource,
                    //
                    worksheetName: null,
                    cellRow: null,
                    cellColumn: null,
                    valueHeader: ValueHeaders.None,
                    value: null,
                    //
                    errorType: ErrorTypes.UnableToUpdateFile,
                    userMessage: string.Format(UserErrorMessages.UnableToUpdateFile, Context.EpplusHelperDataSource.FilePathInUse)
                    );
            }
        }
    }
}