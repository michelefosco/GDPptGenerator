using FilesEditor.Entities;
using FilesEditor.Enums;
using System.IO;

namespace FilesEditor.Steps.BuildPresentation
{
    internal class Step_PredisponiTmpFolder : StepBase
    {
        /// <summary>
        /// Predisposizione della cartella temporanea per la generazione dei file di presentazione.
        /// </summary>
        /// <param name="context"></param>
        public Step_PredisponiTmpFolder(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            predisponiTmpFolder();
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void predisponiTmpFolder()
        {
            // Rimuovo la cartella se già esistente
            CancellaDirectorySeEsiste(Context.TmpFolder);

            // Creo la cartella
            Directory.CreateDirectory(Context.TmpFolder);
        }
    }
}