using System.ComponentModel;

namespace PptGenerator.Enums
{
    public enum TipologiaErrori
    {
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

        [Description("Dato mancante")]
        DatoMancante,

        [Description("Dato non valido")]
        DatoNonValido,

        [Description("Dato non univoco")]
        DatoNonUnivoco
    }
}