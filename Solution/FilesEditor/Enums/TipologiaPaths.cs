using System.ComponentModel;

namespace FilesEditor.Enums
{
    public enum TipologiaPaths
    {




        // Paths database
        [Description("Database folder")]
        DatabaseFolder,

        [Description("Consumption file")]
        ConsumptionFile,

        [Description("Aliases file")]
        AliasesFile,


        // Paths sold
        [Description("Sold folder")]
        SoldFolder,

        [Description("Sold file")]
        SoldFile,


        // Paths output
        [Description("Output folder")]
        OutputFolder,

        [Description("Output file")]
        OutputFile,

        [Description("Debug file")]
        DebugFile
    }
}