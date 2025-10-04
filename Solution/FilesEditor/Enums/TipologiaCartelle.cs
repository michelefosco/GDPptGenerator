using System.ComponentModel;

namespace FilesEditor.Enums
{
    public enum TipologiaCartelle
    {
        [Description("Debug")]
        Debug,

        [Description("Data source Template")]
        DataSource_Template,

        [Description("File di tipo 2")]
        FileDiTipo2,

        [Description("File di tipo 3")]
        FileDiTipo3,

        [Description("File di tipo 4")]
        FileDiTipo4 
    }
}