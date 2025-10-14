using FilesEditor.Entities;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using System.IO;

namespace FilesEditor.Steps.BuildPresentation
{
    internal class Step_PredisponiTmpFolder : StepBase
    {
        public Step_PredisponiTmpFolder(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            predisposiTmpFolder();
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void predisposiTmpFolder()
        {
            // Rimuovo la cartella se già esistente
            if (Directory.Exists(Context.TmpFolder))
            { Directory.Delete(Context.TmpFolder, true); }

            // Creo la cartella
            Directory.CreateDirectory(Context.TmpFolder);
        }
    }
}