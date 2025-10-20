namespace FilesEditor.Constants
{
    public class UserErrorMessages
    {
        public const string UnableToOpenFile = "The selected file could not be opened. Make sure that the file is not already open and that it is in a compatible format.";
        public const string UnableToDeleteFile = "The tool needs to delete the file '{0}' but this is already in use or protected. Please make sure it is deleted.";
        public const string UnableToDeleteFolder = "The tool needs to delete the folder '{0}' but this is already in use or protected. Please make sure it is deleted.";
        public const string UnableToCreateFolder = "The tool needs to create the file '{0}' but there was an error during the attempt of creating it";

        public const string MissingWorksheet = "The worksheet '{0}' is missing";
        public const string MissingValue = "The required value '{0}' is missing";
        public const string MissingHeader = "The file '{0}' does not containt the expected header '{1}' inside the worksheet '{2}'";
        public const string InvalidValue = "Invalid value '{0}'";



        //public const string FoglioMancante = "Nella cartella '{0}' risulta mancare il foglio '{1}'";
        //public const string FoglioIncompleto = "Nella cartella '{0}' il foglio '{1}' risulta incompleto";
        //public const string FileMancante = "Il file atteso non è presente al suo percorso. Verificare che il file esista e ci siano i permessi necessari per la lettura da parte dell'applicazione.";
        //public const string FileGiaEsistente = "Il percorso per il file di output esiste. Sceglierne un altro o eliminare quello già esistente.";
        //public const string CorsoOrarioFornitoreNonValido = "Costo orario non valido per il fornitore con sigla '{0}'";
        //public const string PercorsoFileMancante = "Il percorso del file non può essere vuoto";
    }
}