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
        public bool EvidenziaErroriNelFileDiInput { get; private set; }


        public CreatePresentationsInput(
            string outputFolder,
            string tmpFolder,
            string templatesFolder,
            string fileDebug_FilePath = null,
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
        }
    }
}