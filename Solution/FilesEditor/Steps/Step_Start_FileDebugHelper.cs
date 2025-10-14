using FilesEditor.Entities;
using FilesEditor.Helpers;
using System.IO;
using FilesEditor.Enums;

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
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void start_DebugInfoLogger()
        {
            CancellaFileSeEsiste(Context.FileDebugPath, FileTypes.Debug);
            Context.DebugInfoLogger = new DebugInfoLogger(Context.FileDebugPath, Context.Configurazione.AutoSaveDebugFile);
        }
    }
}