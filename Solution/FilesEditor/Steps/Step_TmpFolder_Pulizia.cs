using FilesEditor.Entities;
using FilesEditor.Enums;

namespace FilesEditor.Steps
{
    internal class Step_TmpFolder_Pulizia : StepBase
    {
        /// <summary>
        /// Predisposizione della cartella temporanea per la generazione dei file di presentazione.
        /// </summary>
        /// <param name="context"></param>
        public Step_TmpFolder_Pulizia(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            Context.DebugInfoLogger.LogStepContext("Step_TmpFolder_Pulizia", Context);
            rimozioneFolder();
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void rimozioneFolder()
        {
            // Rimuovo la cartella se già esistente
            // todo: scommentare
            // CancellaDirectorySeEsiste(Context.TmpFolder);
        }
    }
}