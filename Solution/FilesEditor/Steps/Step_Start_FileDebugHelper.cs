using FilesEditor.Entities;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using System;

namespace FilesEditor.Steps
{
    /// <summary>
    /// Avvio il logging sul file "Debug"
    /// </summary>
    internal class Step_Start_DebugInfoLogger : StepBase
    {
        public override string StepName => "Step_Start_DebugInfoLogger";

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

        public Step_Start_DebugInfoLogger(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            Dtart_DebugInfoLogger();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void Dtart_DebugInfoLogger()
        {
            FilesAndDirectoriesUtilities.CancellaFileSeEsiste(Context.DebugFilePath, FileTypes.Debug);
            Context.SetDebugInfoLogger(new DebugInfoLogger(Context.DebugFilePath, Context.Configurazione.AutoSaveDebugFile));
        }
    }
}