using FilesEditor.Entities;
using FilesEditor.Enums;
using System.IO;

namespace FilesEditor.Steps.BuildPresentation
{
    internal class Step_BackupFile_DataSource :StepBase
    {
        public Step_BackupFile_DataSource(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            Context.DebugInfoLogger.LogStepContext("Step_BackupFile_DataSource", Context);
            backupFile_DataSource();
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void backupFile_DataSource()
        {
            // percors del file DataSource
            Context.DataSourcePath = Path.Combine(Context.SourceFilesFolder, Constants.FileNames.DATA_SOURCE_FILENAME);

            // percorso cartella e file di backup
            var backupFolder = Path.Combine(Context.SourceFilesFolder, Constants.FolderNames.DATASOURCE_FILES_BACKUP_FOLDER);
            CreaDirectorySeNonEsiste(backupFolder); // creo la cartella            
            var backupFilePath = Path.Combine(backupFolder, $"{Path.GetFileNameWithoutExtension(Constants.FileNames.DATA_SOURCE_FILENAME)}_Backup_{System.DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(Constants.FileNames.DATA_SOURCE_FILENAME)}");

            // copio il file
            File.Copy(Context.DataSourcePath, backupFilePath, false);
        }
    }
}