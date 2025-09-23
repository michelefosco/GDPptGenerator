using ReportRefresher.Entities;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Salvataggio del file report nella sua nuova versione
    /// </summary>
    internal class Step_SalvataggioNuovaVersioneFileReport : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            context.InfoFileReport.EPPlusHelper.SaveAs(context.UpdateReportsInput.NewReport_FilePath);
            context.DebugInfoLogger.LogText("Salvataggio nuova versione del report", context.UpdateReportsInput.NewReport_FilePath);

            return null;
        }
    }
}