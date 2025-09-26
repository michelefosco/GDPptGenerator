using FilesEditor.Constants;
using FilesEditor.Entities;
using System.IO;

namespace FilesEditor.Steps
{
    internal class Step_PredisponiTmpFolder : Step_Base
    {
        public Step_PredisponiTmpFolder(StepContext context) : base(context)
        { }

        internal override CreatePresentationsOutput DoSpecificTask()
        {
            predisposiTmpFolder();
            return null; // Step intermedio, non ritorna alcun esito
        }

        private void predisposiTmpFolder()
        {
            // Imposto il percorso della cartella temporanea (sottocartella di OutputFolder)
            Context.TmpFolder = Path.Combine(Context.OutputFolder, FolderNames.TMP_FOLDER_FOR_GENERATED_FILES);

            // Rimuovo la cartella se già esistente
            if (Directory.Exists(Context.TmpFolder))
            { Directory.Delete(Context.TmpFolder, true); }

            // Creo la cartella
            Directory.CreateDirectory(Context.TmpFolder);
        }
    }
}