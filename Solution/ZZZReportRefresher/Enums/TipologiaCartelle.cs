using System.ComponentModel;

namespace ReportRefresher.Enums
{
    public enum TipologiaCartelle
    {
        [Description("Controller")]
        Controller,
        
        [Description("Report input")]
        ReportInput,
        
        [Description("Report output")]
        ReportOutput,
        
        [Description("Debug")]
        Debug
    }
}