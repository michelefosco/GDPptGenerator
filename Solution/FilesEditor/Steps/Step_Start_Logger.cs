using FilesEditor.Entities;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using Serilog;


namespace FilesEditor.Steps
{
    /// <summary>
    /// Avvio il logging sul file 
    /// </summary>
    internal class Step_Start_Logger : StepBase
    {
        internal override string StepName => "Step_Start_Logger";

        public Step_Start_Logger(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            Start_Serilog_Logger();
            Start_DebugInfoLogger();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void Start_Serilog_Logger()
        {
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.File("PptGeneratorGUI.log",
                                rollingInterval: RollingInterval.Day,
                                buffered: true)
                        .CreateLogger();
            Log.Information("Log started");
        }

        private void Start_DebugInfoLogger()
        {
            FilesAndDirectoriesUtilities.CancellaFileSeEsiste(Context.DebugFilePath, FileTypes.Debug);
            Context.SetDebugInfoLogger(new DebugInfoLogger(Context.DebugFilePath, Context.Configurazione.AutoSaveDebugFile));
        }
    }
}