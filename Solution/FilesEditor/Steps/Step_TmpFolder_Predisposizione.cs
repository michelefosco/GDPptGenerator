using FilesEditor.Entities;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using System;
using System.IO;

namespace FilesEditor.Steps
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_TmpFolder_Predisposizione : StepBase
    {
        public override string StepName => "Step_TmpFolder_Predisposizione";

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
        public Step_TmpFolder_Predisposizione(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            PredisponiTmpFolder();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void PredisponiTmpFolder()
        {
            // Rimuovo la cartella se già esistente
            FilesAndDirectoriesUtilities.CancellaDirectorySeEsiste(Context.TmpFolder);

            // Creo la cartella
            Directory.CreateDirectory(Context.TmpFolder);
        }
    }
}