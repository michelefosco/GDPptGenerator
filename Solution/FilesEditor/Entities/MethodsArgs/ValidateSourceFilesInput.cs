using FilesEditor.Entities.MethodsArgs;
using System;

namespace FilesEditor.Entities
{
    public class ValidateSourceFilesInput : UserInterfaceInputBase
    {
        public string FileBudgetPath { get; private set; }
        public string FileForecastPath { get; private set; }
        public string FileSuperDettagliPath { get; private set; }
        public string FileRunRatePath { get; private set; }

        public ValidateSourceFilesInput(
                string dataSourceFilePath,
                string destinationFolder,
                string tmpFolder,
                string debugFilePath,
                //
                string fileBudgetPath,
                string fileForecastPath,
                string fileSuperDettagliPath,
                string fileRunRatePath
            )
        {
            // Properties from the base class
            if (string.IsNullOrWhiteSpace(dataSourceFilePath))
                throw new ArgumentNullException(nameof(dataSourceFilePath));
            if (string.IsNullOrWhiteSpace(destinationFolder))
                throw new ArgumentNullException(nameof(destinationFolder));
            if (string.IsNullOrWhiteSpace(tmpFolder))
                throw new ArgumentNullException(nameof(tmpFolder));
            if (string.IsNullOrWhiteSpace(debugFilePath))
                throw new ArgumentNullException(nameof(DebugFilePath));
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
            base.DataSourceFilePath = dataSourceFilePath;
            base.DestinationFolder = destinationFolder;
            base.TmpFolder = tmpFolder;
            base.DebugFilePath = debugFilePath;

            // Properties of the derived class 
            FileBudgetPath = fileBudgetPath;
            FileForecastPath = fileForecastPath;
            FileSuperDettagliPath = fileSuperDettagliPath;
            FileRunRatePath = fileRunRatePath;
        }
    }
}