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
        internal override string StepName => "Step_Start_DebugInfoLogger";

        public Step_Start_DebugInfoLogger(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            Start_DebugInfoLogger();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void Start_DebugInfoLogger()
        {
            FilesAndDirectoriesUtilities.CancellaFileSeEsiste(Context.DebugFilePath, FileTypes.Debug);
            Context.SetDebugInfoLogger(new DebugInfoLogger(Context.DebugFilePath, Context.Configurazione.AutoSaveDebugFile));
        }
    }
}