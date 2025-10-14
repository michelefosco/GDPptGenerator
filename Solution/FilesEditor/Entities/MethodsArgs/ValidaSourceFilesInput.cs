using FilesEditor.Entities.MethodsArgs;
using System;

namespace FilesEditor.Entities
{
    public class ValidaSourceFilesInput : UserInterfaceInputBase
    {
        public string FileBudgetPath { get; private set; }
        public string FileForecastPath { get; private set; }
        public string FileSuperDettagliPath { get; private set; }
        public string FileRunRatePath { get; private set; }

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
            // Properties from the base class
            if (string.IsNullOrWhiteSpace(sourceFilesFolderPath))
                throw new ArgumentNullException(nameof(sourceFilesFolderPath));
            if (string.IsNullOrWhiteSpace(destinationFolder))
                throw new ArgumentNullException(nameof(destinationFolder));
            if (string.IsNullOrWhiteSpace(tmpFolder))
                throw new ArgumentNullException(nameof(tmpFolder));
            if (string.IsNullOrWhiteSpace(fileDebugPath))
                throw new ArgumentNullException(nameof(fileDebugPath));
            // Properties of the derived class 
            if (string.IsNullOrWhiteSpace(fileBudgetPath))
                throw new ArgumentNullException(nameof(fileBudgetPath));
            if (string.IsNullOrWhiteSpace(fileForecastPath))
                throw new ArgumentNullException(nameof(fileForecastPath));
            if (string.IsNullOrWhiteSpace(fileSuperDettagliPath))
                throw new ArgumentNullException(nameof(fileSuperDettagliPath));
            if (string.IsNullOrWhiteSpace(fileRunRatePath))
                throw new ArgumentNullException(nameof(fileRunRatePath));

            // Properties from the base class
            base.SourceFilesFolder = sourceFilesFolderPath;
            base.DestinationFolder = destinationFolder;
            base.TmpFolder = tmpFolder;
            base.FileDebugPath = fileDebugPath;

            // Properties of the derived class 
            FileBudgetPath = fileBudgetPath;
            FileForecastPath = fileForecastPath;
            FileSuperDettagliPath = fileSuperDettagliPath;
            FileRunRatePath = fileRunRatePath;
        }
    }
}