using ReportRefresher.Entities;
using ReportRefresher.Enums;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_EsitoFinaleSuccess : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            context.UpdateReportsOutput.SettaEsitoFinale(EsitiFinali.Success);
            ChiusuraFileDebug(context);
            
            return context.UpdateReportsOutput;
        }
    }
}