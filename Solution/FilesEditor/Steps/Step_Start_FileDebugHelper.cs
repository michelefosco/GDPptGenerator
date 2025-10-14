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
            inizializzaDebugInfoLogger();
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void inizializzaDebugInfoLogger()
        {
            if (File.Exists(Context.FileDebugPath))
            {
                File.Delete(Context.FileDebugPath);
            }
            Context.DebugInfoLogger = new DebugInfoLogger(Context.FileDebugPath, Context.Configurazione.AutoSaveDebugFile);
        }
    }
}