using FilesEditor.Entities;
using FilesEditor.Enums;
using System;

namespace FilesEditor.Steps
{
    /// <summary>
    /// Predisposizione della cartella temporanea per la generazione dei file di presentazione.
    /// </summary>
    /// <param name="context"></param>
    internal class Step_TmpFolder_Pulizia : StepBase
    {
        public override string StepName => "Step_TmpFolder_Pulizia";

        internal override void BeforeTask()
        {
            Context.DebugInfoLogger.LogStepContext(StepName, Context);
        }

        internal override void ManageInfoAboutPerformedStepTask(TimeSpan timeSpent)
        {
            Context.DebugInfoLogger.LogPerformance(StepName, timeSpent);
        }

        internal override void AfterTask()
        {
            Context.DebugInfoLogger.LogStepContext(StepName, Context);
        }

        public Step_TmpFolder_Pulizia(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoStepTask()
        {
            RimozioneFolder();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void RimozioneFolder()
        {
            // Rimuovo la cartella se già esistente
            // todo: scommentare
            // CancellaDirectorySeEsiste(Context.TmpFolder);
        }
    }
}