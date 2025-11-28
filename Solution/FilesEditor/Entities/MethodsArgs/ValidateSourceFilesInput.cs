using FilesEditor.Entities.MethodsArgs;
using System;

namespace FilesEditor.Entities
{
    public class ValidateSourceFilesInput : UserInterfaceInputBase
    {
        public ValidateSourceFilesInput(
                string dataSourceFilePath,
                string destinationFolder,
                string tmpFolder,
                string debugFilePath,
                //
                string fileBudgetPath,
                string fileForecastPath,
                string fileSuperDettagliPath,
                string fileRunRatePath,
                string fileCN43NPath
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
            // 
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
            //
            base.FileBudgetPath = fileBudgetPath;
            base.FileForecastPath = fileForecastPath;
            base.FileSuperDettagliPath = fileSuperDettagliPath;
            base.FileRunRatePath = fileRunRatePath;
            base.FileCN43NPath = fileCN43NPath;            
        }
    }
}