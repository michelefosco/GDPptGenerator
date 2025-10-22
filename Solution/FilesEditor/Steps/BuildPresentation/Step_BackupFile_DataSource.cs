using FilesEditor.Entities;
using FilesEditor.Enums;
using System.IO;

namespace FilesEditor.Steps.BuildPresentation
{
    internal class Step_BackupFile_DataSource : StepBase
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
            // percorso cartella di backup
            var sourceFilesFolder = Path.GetDirectoryName(Context.DataSourceFilePath);
            var backupFolder = Path.Combine(sourceFilesFolder, Constants.FolderNames.DATASOURCE_FILES_BACKUP_FOLDER);
            CreaDirectorySeNonEsiste(backupFolder); // creo la cartella            

            // percorso file di backup
            var backupFilePath = Path.Combine(backupFolder, $"{Path.GetFileNameWithoutExtension(Constants.FileNames.DATASOURCE_FILENAME)}_Backup_{System.DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(Constants.FileNames.DATASOURCE_FILENAME)}");

            // copio il file
            File.Copy(Context.DataSourceFilePath, backupFilePath, false);
        }
    }
}