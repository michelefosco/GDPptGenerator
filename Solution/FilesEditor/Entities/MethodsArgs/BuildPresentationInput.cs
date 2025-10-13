using System;

namespace FilesEditor.Entities.MethodsArgs
{
    public class BuildPresentationInput
    {
        public string OutputFolder { get; private set; }
        public string TmpFolder { get; private set; }
        public string TemplatesFolder { get; private set; }
        public string FileDebug_FilePath { get; private set; }

        //public bool ReplaceAllData_FileBudget { get; private set; }
        //public bool ReplaceAllData_FileForecast { get; private set; }
        public bool ReplaceAllData_FileSuperDettagli { get; private set; }

        public DateTime PeriodDate { get; private set; }

        //public bool EvidenziaErroriNelFileDiInput { get; private set; }


        public BuildPresentationInput(
            string outputFolder,
            string tmpFolder,
            string templatesFolder,
            string fileDebug_FilePath,
            //bool replaceAllData_FileBudget,
            //bool replaceAllData_FileForecast,
            bool replaceAllData_FileSuperDettagli,
            DateTime periodDate
            //bool evidenziaErroriNelFileDiInput = false
            )
        {
            if (string.IsNullOrWhiteSpace(outputFolder))
                throw new ArgumentNullException(nameof(outputFolder));
            if (string.IsNullOrWhiteSpace(tmpFolder))
                throw new ArgumentNullException(nameof(tmpFolder));
            if (string.IsNullOrWhiteSpace(templatesFolder))
                throw new ArgumentNullException(nameof(templatesFolder));

            OutputFolder = outputFolder;
            TmpFolder = tmpFolder;
            TemplatesFolder = templatesFolder;
            FileDebug_FilePath = fileDebug_FilePath;

            //ReplaceAllData_FileBudget = replaceAllData_FileBudget;
            //ReplaceAllData_FileForecast = replaceAllData_FileForecast;
            ReplaceAllData_FileSuperDettagli = replaceAllData_FileSuperDettagli;

            PeriodDate = periodDate;
        }
    }
}