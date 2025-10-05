using System.ComponentModel;

namespace FilesEditor.Enums
{
    public enum FileTypes
    {
        None = 0,

        [Description("Debug")]
        Debug,

        [Description("File Excel DataSource")]
        DataSource,

        [Description("File Excel DataSource Template")]
        DataSource_Template,

        [Description("File Excel Super Dettagli")]
        SuperDettagli,

        [Description("File Excel Budget")]
        Budget,

        [Description("File Excel Forecast")]
        Forecast,

        [Description("File Excel Ran Rate")]
        RanRate
    }
}