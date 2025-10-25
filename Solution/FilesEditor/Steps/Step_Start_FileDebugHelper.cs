using FilesEditor.Entities;
using FilesEditor.Enums;
using FilesEditor.Helpers;

namespace FilesEditor.Steps
{
    /// <summary>
    /// Avvio il logging sul file "Debug"
    /// </summary>
    internal class Step_Start_DebugInfoLogger : StepBase
    {
        public Step_Start_DebugInfoLogger(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            start_DebugInfoLogger();

            Context.DebugInfoLogger.LogStepContext("Step_Start_DebugInfoLogger", Context);

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void start_DebugInfoLogger()
        {
            FilesAndDirectoriesUtilities.CancellaFileSeEsiste(Context.DebugFilePath, FileTypes.Debug);
            Context.SetDebugInfoLogger(new DebugInfoLogger(Context.DebugFilePath, Context.Configurazione.AutoSaveDebugFile));
        }
    }
}