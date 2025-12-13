using FilesEditor.Entities;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using System;

namespace FilesEditor.Steps
{
    /// <summary>
    /// Predisposizione della cartella temporanea per la generazione dei file di presentazione.
    /// </summary>
    /// <param name="context"></param>
    internal class Step_TmpFolder_Pulizia : StepBase
    {
        internal override string StepName => "Step_TmpFolder_Pulizia";

        public Step_TmpFolder_Pulizia(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            RimozioneFolder();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void RimozioneFolder()
        {
            // Rimuovo la cartella se già esistente
            FilesAndDirectoriesUtilities.CancellaDirectorySeEsiste(Context.TmpFolder);
        }
    }
}