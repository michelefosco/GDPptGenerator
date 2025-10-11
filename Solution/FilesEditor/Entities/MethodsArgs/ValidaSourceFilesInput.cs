using System;

namespace FilesEditor.Entities
{
    public class ValidaSourceFilesInput
    {
        public string DestinationFolder { get; private set; }
        public string SourceFilesFolderPath { get; private set; }
        public string FileBudgetPath { get; private set; }
        public string FileForecastPath { get; private set; }
        public string FileSuperDettagliPath { get; private set; }
        public string FileRunRatePath { get; private set; }

        //public string DestinationFolderPath { get; private set; }
        //public string OutputFolder { get; private set; }
        //public string TmpFolder { get; private set; }
        //public bool EvidenziaErroriNelFileDiInput { get; private set; }


        public ValidaSourceFilesInput(
                 string destinationFolder,
                 string sourceFilesFolderPath,
                 string fileBudgetPath,
                 string fileForecastPath,
                 string fileSuperDettagliPath,
                 string fileRunRatePath
            )
        {
            if (string.IsNullOrWhiteSpace(destinationFolder))
                throw new ArgumentNullException(nameof(destinationFolder));
            if (string.IsNullOrWhiteSpace(sourceFilesFolderPath))
                throw new ArgumentNullException(nameof(sourceFilesFolderPath));
            if (string.IsNullOrWhiteSpace(fileBudgetPath))
                throw new ArgumentNullException(nameof(fileBudgetPath));
            if (string.IsNullOrWhiteSpace(fileForecastPath))
                throw new ArgumentNullException(nameof(fileForecastPath));
            if (string.IsNullOrWhiteSpace(fileSuperDettagliPath))
                throw new ArgumentNullException(nameof(fileSuperDettagliPath));
            if (string.IsNullOrWhiteSpace(fileRunRatePath))
                throw new ArgumentNullException(nameof(fileRunRatePath));

            DestinationFolder = destinationFolder;
            SourceFilesFolderPath = sourceFilesFolderPath;
            FileBudgetPath = fileBudgetPath;
            FileForecastPath = fileForecastPath;
            FileSuperDettagliPath = fileSuperDettagliPath;
            FileRunRatePath = fileRunRatePath;
        }
    }
}