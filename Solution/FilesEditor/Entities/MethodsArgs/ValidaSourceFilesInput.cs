using FilesEditor.Entities.MethodsArgs;
using System;

namespace FilesEditor.Entities
{
    public class ValidaSourceFilesInput : UserInterfaceInputBase
    {
        //public string DestinationFolder { get; private set; }
        //public string TmpFolder { get; private set; }
        //public string SourceFilesFolderPath { get; private set; }


        public string FileBudgetPath { get; private set; }
        public string FileForecastPath { get; private set; }
        public string FileSuperDettagliPath { get; private set; }
        public string FileRunRatePath { get; private set; }

        //public string DestinationFolderPath { get; private set; }
        //public string OutputFolder { get; private set; }
        //public string TmpFolder { get; private set; }
        //public bool EvidenziaErroriNelFileDiInput { get; private set; }


        public ValidaSourceFilesInput(
                string sourceFilesFolderPath,
                string destinationFolder,
                string tmpFolder,
                string fileDebugPath,
                //
                string fileBudgetPath,
                string fileForecastPath,
                string fileSuperDettagliPath,
                string fileRunRatePath
            )
        {
            if (string.IsNullOrWhiteSpace(sourceFilesFolderPath))
                throw new ArgumentNullException(nameof(sourceFilesFolderPath));
            if (string.IsNullOrWhiteSpace(destinationFolder))
                throw new ArgumentNullException(nameof(destinationFolder));
            if (string.IsNullOrWhiteSpace(tmpFolder))
                throw new ArgumentNullException(nameof(tmpFolder));
            if (string.IsNullOrWhiteSpace(fileDebugPath))
                throw new ArgumentNullException(nameof(fileDebugPath));
            //
            if (string.IsNullOrWhiteSpace(fileBudgetPath))
                throw new ArgumentNullException(nameof(fileBudgetPath));
            if (string.IsNullOrWhiteSpace(fileForecastPath))
                throw new ArgumentNullException(nameof(fileForecastPath));
            if (string.IsNullOrWhiteSpace(fileSuperDettagliPath))
                throw new ArgumentNullException(nameof(fileSuperDettagliPath));
            if (string.IsNullOrWhiteSpace(fileRunRatePath))
                throw new ArgumentNullException(nameof(fileRunRatePath));
            
            // Properies base class
            base.SourceFilesFolder = sourceFilesFolderPath;
            base.DestinationFolder = destinationFolder;
            base.TmpFolder = tmpFolder;
            base.FileDebugPath = fileDebugPath;

            //
            FileBudgetPath = fileBudgetPath;
            FileForecastPath = fileForecastPath;
            FileSuperDettagliPath = fileSuperDettagliPath;
            FileRunRatePath = fileRunRatePath;
        }
    }
}