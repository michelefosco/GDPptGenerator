using System.ComponentModel;

namespace FilesEditor.Enums
{
    public enum ValueHeaders
    {
        None = 0,

        [Description("Table name")]
        TableName,

        [Description("Presentation file name")]
        PresentationFileName,

        [Description("Slide title")]
        SlideTitle,

        [Description("Slide content")]
        SlideContent,

        [Description("Slide layout")]
        SlideLayout,
    }
}