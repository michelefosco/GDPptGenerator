using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using System;
using System.IO;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// Avvio il DataSource EPPlusHelper per assicurarmi che il file sia editabile
    /// </summary>
    internal class Step_VerificaEditabilita_DataSource_File : StepBase
    {
        public override string StepName => "Step_VerificaEditabilita_DataSource_File";

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
        public Step_VerificaEditabilita_DataSource_File(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            AttemptToOpenDataSourceFile();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void AttemptToOpenDataSourceFile()
        {
            try
            {
                using (FileStream stream = new FileStream(
                    Context.DataSourceEPPlusHelper.FilePathInUse,
                    FileMode.Open,
                    FileAccess.ReadWrite,
                    FileShare.None)) // <– exclusive lock
                {
                    var canAccessFile = stream.CanWrite;
                }
            }
            catch (IOException)
            {
                throw new ManagedException(
                    filePath: Context.DataSourceEPPlusHelper.FilePathInUse,
                    fileType: FileTypes.DataSource,
                    //
                    worksheetName: null,
                    cellRow: null,
                    cellColumn: null,
                    valueHeader: ValueHeaders.None,
                    value: null,
                    //
                    errorType: ErrorTypes.UnableToUpdateFile,
                    userMessage: string.Format(UserErrorMessages.UnableToUpdateFile, Context.DataSourceEPPlusHelper.FilePathInUse)
                    );
            }
        }
    }
}