using System.ComponentModel;

namespace FilesEditor.Enums
{
    public enum ErrorTypes
    {
        None = 0,

        [Description("Unable to open a file")]
        UnableToOpenFile,

        [Description("Unable to delete a file")]
        UnableToDeleteFile,

        [Description("Unable to delete a folder")]
        UnableToDeleteFolder,

        [Description("Unable to create a folder")]
        UnableToCreateFolder,


        [Description("An expected worksheet is missing")]
        MissingWorksheet,

        [Description("Missing value")]
        MissingValue,




        [Description("Impossibile creare il file")]
        ImpossibileCreareFile,

        [Description("File già esistente")]
        FileGiaEsistente,

        [Description("File inesistente")]
        FileMancante,

        [Description("Formato file errato")]
        FormatoFileErrato,

        [Description("Foglio mancante")]
        FoglioMancante,

        [Description("Foglio incompleto")]
        FoglioIncompleto,

        [Description("Dato non valido")]
        DatoNonValido,

        [Description("Dato non univoco")]
        DatoNonUnivoco,

        [Description("Errore in una formula")]
        Formula,
        UnhandledException
    }
}