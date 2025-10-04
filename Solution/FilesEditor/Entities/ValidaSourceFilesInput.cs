using System;

namespace FilesEditor.Entities
{
    public class ValidaSourceFilesInput
    {
        public string TemplatesFolder { get; private set; }
        //public string OutputFolder { get; private set; }
        //public string TmpFolder { get; private set; }
        //public bool EvidenziaErroriNelFileDiInput { get; private set; }
        //public string FileDebug_FilePath { get; private set; }

        public ValidaSourceFilesInput(
            string templateFolder
            //string outputFolder,
            //string tmpFolder,            
            //string fileDebug_FilePath = null,
            //bool evidenziaErroriNelFileDiInput = false
            )
        {
            if (string.IsNullOrWhiteSpace(templateFolder))
                throw new ArgumentNullException(nameof(templateFolder));
            //if (string.IsNullOrWhiteSpace(outputFolder))
            //    throw new ArgumentNullException(nameof(outputFolder));
            //if (string.IsNullOrWhiteSpace(tmpFolder))
            //    throw new ArgumentNullException(nameof(tmpFolder));



            TemplatesFolder = templateFolder;
            //OutputFolder = outputFolder;
            //TmpFolder = tmpFolder;
            //FileDebug_FilePath = fileDebug_FilePath;
            //EvidenziaErroriNelFileDiInput = evidenziaErroriNelFileDiInput;
        }
    }
}