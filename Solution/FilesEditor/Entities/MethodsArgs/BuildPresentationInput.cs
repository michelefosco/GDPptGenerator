using System;
using System.Collections.Generic;

namespace FilesEditor.Entities.MethodsArgs
{
    public class BuildPresentationInput : UserInterfaceInputBase
    {
        public bool ReplaceAllData_FileSuperDettagli { get; private set; }
        public DateTime PeriodDate { get; private set; }
        public List<InputDataFilters_Item> Applicablefilters { get; private set; }

        public BuildPresentationInput(
                string sourceFilesFolderPath,
                string destinationFolder,
                string tmpFolder,
                string fileDebugPath,
                //
                bool replaceAllData_FileSuperDettagli,
                DateTime periodDate,
                List<InputDataFilters_Item> applicablefilters
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
            //...


            // Properties from the base class
            base.SourceFilesFolder = sourceFilesFolderPath;
            base.DestinationFolder = destinationFolder;
            base.TmpFolder = tmpFolder;
            base.FileDebugPath = fileDebugPath;

            // Properties of the derived class 
            ReplaceAllData_FileSuperDettagli = replaceAllData_FileSuperDettagli;
            PeriodDate = periodDate;
            //todo utilizzare
            Applicablefilters = applicablefilters ?? new List<InputDataFilters_Item>();
        }
    }
}