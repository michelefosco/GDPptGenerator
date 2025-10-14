using System.ComponentModel;

namespace FilesEditor.Enums
{
    public enum FileTypes
    {
        Undefined,

        Directory,

        [Description("Debug")]
        Debug,

        [Description("File Excel DataSource")]
        DataSource,

        [Description("File Excel (input) Super Dettagli")]
        SuperDettagli,

        [Description("File Excel (input) Budget")]
        Budget,

        [Description("File Excel (input) Forecast")]
        Forecast,

        [Description("File Excel (input) Run Rate")]
        RunRate,

        [Description("File PowerPoint Presentation template")]
        PresentationTemplate,

        [Description("File PowerPoint Presentation output")]
        PresentationOutput
    }
}