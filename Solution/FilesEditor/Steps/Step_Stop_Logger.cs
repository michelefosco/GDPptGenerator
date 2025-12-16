using FilesEditor.Entities;
using FilesEditor.Enums;

namespace FilesEditor.Steps
{
    /// <summary>
    /// Avvio il logging sul file 
    /// </summary>
    internal class Step_Stop_Logger : StepBase
    {
        internal override string StepName => "Step_Stop_Logger";

        public Step_Stop_Logger(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            Stop_Logger();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void Stop_Logger()
        {
            Context.DebugInfoLogger.Beautify();

            Serilog.Log.Information("CloseAndFlush...");
            Serilog.Log.CloseAndFlush();
        }
    }
}