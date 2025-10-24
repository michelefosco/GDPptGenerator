using System.ComponentModel;

namespace FilesEditor.Enums
{
    public enum ErrorTypes
    {
        None = 0,

        // files and folders errors
        [Description("Unable to open a file")]
        UnableToOpenFile,

        [Description("Unable to update a file")]
        UnableToUpdateFile,

        [Description("Unable to delete a file")]
        UnableToDeleteFile,

        [Description("Unable to delete a folder")]
        UnableToDeleteFolder,

        [Description("Unable to create a folder")]
        UnableToCreateFolder,

        [Description("File already exists")]
        FileAlreadyExists,

        [Description("Unable to create the file")]
        UnableToCreateFile,


        // Working with Excel files errors
        [Description("An expected worksheet is missing")]
        MissingWorksheet,



        // Data errors
        [Description("Missing value")]
        MissingValue,

        [Description("Duplicate value")]
        DuplicateValue,

        [Description("Invalid value")]
        InvalidValue,

        [Description("Missing header")]
        MissingHeader,



        // not used yet

        //[Description("File inesistente")]
        //FileMancante,

        //[Description("Formato file errato")]
        //FormatoFileErrato,

        //[Description("Foglio mancante")]
        //FoglioMancante,

        //[Description("Foglio incompleto")]
        //FoglioIncompleto,

        //[Description("Errore in una formula")]
        //Formula,

        UnhandledException
    }
}