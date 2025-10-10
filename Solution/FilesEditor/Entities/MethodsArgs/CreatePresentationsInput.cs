using System;
using System.Collections.Generic;

namespace FilesEditor.Entities.MethodsArgs
{
    public class CreatePresentationsInput
    {
        public string OutputFolder { get; private set; }
        public string TmpFolder { get; private set; }
        public string TemplatesFolder { get; private set; }
        public string FileDebug_FilePath { get; private set; }

        public bool ReplaceAllData_FileBudget { get; private set; }
        public bool ReplaceAllData_FileForecast { get; private set; }
        public bool ReplaceAllData_FileRunRate { get; private set; }
        public bool ReplaceAllData_FileSuperDettagli { get; private set; }

        public bool EvidenziaErroriNelFileDiInput { get; private set; }


        public CreatePresentationsInput(
            string outputFolder,
            string tmpFolder,
            string templatesFolder,
            string fileDebug_FilePath,
            bool replaceAllData_FileBudget,
            bool replaceAllData_FileForecast,
            bool replaceAllData_FileRunRate,
            bool replaceAllData_FileSuperDettagli,

        bool evidenziaErroriNelFileDiInput = false)
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
            EvidenziaErroriNelFileDiInput = evidenziaErroriNelFileDiInput;

            ReplaceAllData_FileBudget = replaceAllData_FileBudget;
            ReplaceAllData_FileForecast = replaceAllData_FileForecast;
            ReplaceAllData_FileRunRate = replaceAllData_FileRunRate;
            ReplaceAllData_FileSuperDettagli = replaceAllData_FileSuperDettagli;
        }
    }
}