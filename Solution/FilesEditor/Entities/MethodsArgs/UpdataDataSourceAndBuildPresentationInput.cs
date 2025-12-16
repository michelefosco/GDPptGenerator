using System;
using System.Collections.Generic;

namespace FilesEditor.Entities.MethodsArgs
{
    public class UpdataDataSourceAndBuildPresentationInput : UserInterfaceInputBase
    {
        public string PowerPointTemplateFilePath { get; private set; }
        public bool AppendCurrentYear_FileSuperDettagli { get; private set; }
        public DateTime PeriodDate { get; private set; }
        public List<InputDataFilters_Item> ApplicableFilters { get; private set; }

        public bool BuildPresentationOnly { get; private set; }

        public UpdataDataSourceAndBuildPresentationInput(
                string dataSourceFilePath,
                string destinationFolder,
                string tmpFolder,
                string debugFilePath,
                //
                string fileBudgetPath,
                string fileForecastPath,
                string fileSuperDettagliPath,
                string fileRunRatePath,
                string fileCN43NPath,
                //
                string powerPointTemplateFilePath,
                bool appendCurrentYear_FileSuperDettagli,
                DateTime periodDate,
                List<InputDataFilters_Item> applicablefilters,
                bool buildPresentationOnly
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

            // Properties of the derived class 
            PowerPointTemplateFilePath = powerPointTemplateFilePath;
            AppendCurrentYear_FileSuperDettagli = appendCurrentYear_FileSuperDettagli;
            PeriodDate = periodDate;
            ApplicableFilters = applicablefilters ?? new List<InputDataFilters_Item>();
            BuildPresentationOnly = buildPresentationOnly;
        }
    }
}