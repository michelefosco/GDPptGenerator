using FilesEditor.Entities;
using FilesEditor.Enums;
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
            Start_Logger();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void Start_Logger()
        {
            //Context.Logger = new LoggerConfiguration()
            //    .MinimumLevel.Information()
            //    .WriteTo.File(
            //        "PptGeneratorGUI.log",
            //        rollingInterval: RollingInterval.Day,
            //        buffered: true)
            //    .CreateLogger();
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.File("PptGeneratorGUI.log",
                                rollingInterval: RollingInterval.Day,
                                buffered: true)
                        .CreateLogger();
            Log.Information("Log started");
        }
    }
}