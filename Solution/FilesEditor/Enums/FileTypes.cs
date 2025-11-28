using System.ComponentModel;

namespace FilesEditor.Enums
{
    public enum FileTypes
    {
        Undefined,

        Directory,

        [Description("Debug")]
        Debug,

        [Description("DataSource Excel file")]
        DataSource,

        [Description("Super Dettagli Excel input file")]
        SuperDettagli,

        [Description("Budget Excel input file")]
        Budget,

        [Description("Forecast Excel input file")]
        Forecast,

        [Description("Run Rate Excel input file")]
        RunRate,

        [Description("CN43N Excel input file")]
        CN43N,        

        [Description("PowerPoint presentation template fuke")]
        PresentationTemplate,

        [Description("PowerPoint Presentation output file")]
        PresentationOutput
    }
}