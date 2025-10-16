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
                string dataSourceFilePath,
                string destinationFolder,
                string tmpFolder,
                string debugFilePath,
                //
                bool replaceAllData_FileSuperDettagli,
                DateTime periodDate,
                List<InputDataFilters_Item> applicablefilters
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
                throw new ArgumentNullException(nameof(debugFilePath));
            // Properties of the derived class 
            //...


            // Properties from the base class
            base.DataSourceFilePath = dataSourceFilePath;
            base.DestinationFolder = destinationFolder;
            base.TmpFolder = tmpFolder;
            base.DebugFilePath = debugFilePath;

            // Properties of the derived class 
            ReplaceAllData_FileSuperDettagli = replaceAllData_FileSuperDettagli;
            PeriodDate = periodDate;
            //todo utilizzare
            Applicablefilters = applicablefilters ?? new List<InputDataFilters_Item>();
        }
    }
}