using System;
using System.Collections.Generic;

namespace FilesEditor.Entities
{
    public class CreatePresentationsInput
    {
        #region Input obbligatorio
        public string OutputFolder { get; private set; }
        public string ConfigurationFolder { get; private set; }
        
        //public string FileReport_FilePath { get; private set; }
        //public string NewReport_FilePath { get; private set; }
        #endregion

        #region Input con valori di default
        public bool EvidenziaErroriNelFileDiInput { get; private set; }
        public string FileDebug_FilePath { get; private set; }
        #endregion

        #region Input opzionale
        #endregion




        public CreatePresentationsInput(
            string outputFolder,
            string configurationFolder,
            //string newReport_FilePath,
            string fileDebug_FilePath = null,
            bool evidenziaErroriNelFileDiInput = false)
        {
            if (string.IsNullOrWhiteSpace(outputFolder))
                throw new ArgumentNullException(nameof(outputFolder));
            if (string.IsNullOrWhiteSpace(configurationFolder))
                throw new ArgumentNullException(nameof(configurationFolder));
            //if (string.IsNullOrWhiteSpace(fileReport_FilePath))
            //    throw new ArgumentNullException(nameof(fileReport_FilePath));
            //if (string.IsNullOrWhiteSpace(newReport_FilePath))
            //    throw new ArgumentNullException(nameof(newReport_FilePath));

            OutputFolder = outputFolder;
            ConfigurationFolder = configurationFolder;
            //FileReport_FilePath = fileReport_FilePath;
            //NewReport_FilePath = newReport_FilePath;
            FileDebug_FilePath = fileDebug_FilePath;
            EvidenziaErroriNelFileDiInput = evidenziaErroriNelFileDiInput;
        }
    }
}