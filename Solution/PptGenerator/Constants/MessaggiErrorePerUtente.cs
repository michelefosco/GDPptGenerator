namespace PptGenerator.Constants
{
    public class MessaggiErrorePerUtente
    {
        public const string FoglioMancante = "Nella cartella '{0}' risulta mancare il foglio '{1}'";
        public const string FoglioIncompleto = "Nella cartella '{0}' il foglio '{1}' risulta incompleto";
        public const string FileMancante = "Il file atteso non è presente al suo percorso. Verificare che il file esista e ci siano i permessi necessari per la lettura da parte dell'applicazione.";
        public const string FileGiaEsistente = "Il percorso per il file di output esiste. Sceglierne un altro o eliminare quello già esistente.";
        public const string FileImpossibileDaAprire = "Non è stato possibile aprire il file selezionato.\nVerificare che il file non sia già aperto e che abbia un formato compatibile";
        public const string CorsoOrarioFornitoreNonValido = "Costo orario non valido per il fornitore con sigla '{0}'";
        public const string PercorsoFileMancante = "Il percorso del file non può essere vuoto";
    }
}