using FilesEditor.Entities;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using System.IO;
using System.IO.Compression;


namespace FilesEditor.Steps.BuildPresentation
{
    internal class Step_BackupFile_DataSource : StepBase
    {
        public Step_BackupFile_DataSource(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            Context.DebugInfoLogger.LogStepContext("Step_BackupFile_DataSource", Context);
            BackupFile_DataSource();
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void BackupFile_DataSource()
        {
            // predisponco la cartella per il backup
            var sourceFilesFolder = Path.GetDirectoryName(Context.DataSourceFilePath);
            var backupFolder = Path.Combine(sourceFilesFolder, Constants.FolderNames.DATASOURCE_FILES_BACKUP_FOLDER);
            FilesAndDirectoriesUtilities.CreaDirectorySeNonEsiste(backupFolder); // creo la cartella            


            if (Context.Configurazione.ZipBackupFile)
            {
                // Zip file destinazione
                var backupZipFilePath = Path.Combine(backupFolder, $"{Path.GetFileNameWithoutExtension(Constants.FileNames.DATASOURCE_FILENAME)}_Backup_{System.DateTime.Now:yyyyMMdd_HHmmss}.zip");

                // Creazioen dell'archivio zip
                using (FileStream zipToOpen = new FileStream(backupZipFilePath, FileMode.Create))
                {
                    using (var archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                    {
                        archive.CreateEntryFromFile(Context.DataSourceFilePath, Path.GetFileName(Context.DataSourceFilePath), CompressionLevel.Optimal);
                    }
                }
            }
            else
            {
                // percorso file di backup
                var backupFilePath = Path.Combine(backupFolder, $"{Path.GetFileNameWithoutExtension(Constants.FileNames.DATASOURCE_FILENAME)}_Backup_{System.DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(Constants.FileNames.DATASOURCE_FILENAME)}");

                // copio il file
                File.Copy(Context.DataSourceFilePath, backupFilePath, false);
            }
        }
    }
}