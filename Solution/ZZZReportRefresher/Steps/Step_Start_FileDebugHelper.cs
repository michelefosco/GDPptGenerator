using ReportRefresher.Entities;
using ReportRefresher.Helpers;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Avvio il logging sul file "Debug"
    /// </summary>
    internal class Step_Start_FileDebugHelper : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            context.DebugInfoLogger = new FileDebugHelper(context.UpdateReportsInput.FileDebug_FilePath, context.Configurazione.AutoSaveDebugFile);
            context.DebugInfoLogger.LogUpdateReportsInput(context.UpdateReportsInput);

            return null;
        }
    }
}