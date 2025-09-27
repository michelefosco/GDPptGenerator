using FilesEditor.Entities;
using FilesEditor.Helpers;
using FilesEditor.Steps;
using System.IO;


namespace FilesEditor.Steps
{
    /// <summary>
    /// Avvio il logging sul file "Debug"
    /// </summary>
    internal class Step_Start_FileDebugHelper : Step_Base
    {

        public Step_Start_FileDebugHelper(StepContext context) : base(context)
        { }

        internal override CreatePresentationsOutput DoSpecificTask()
        {
            inizializzaDebugInfoLogger();
            return null; // Step intermedio, non ritorna alcun esito
        }

        private void inizializzaDebugInfoLogger()
        {
            if (File.Exists(Context.CreatePresentationsInput.FileDebug_FilePath))
            {
                File.Delete(Context.CreatePresentationsInput.FileDebug_FilePath);
            }
            Context.DebugInfoLogger = new FileDebugHelper(Context.CreatePresentationsInput.FileDebug_FilePath, Context.Configurazione.AutoSaveDebugFile);
            Context.DebugInfoLogger.LogCreatePresentationsInput(Context.CreatePresentationsInput);
        }
    }
}