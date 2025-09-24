using System;

namespace FilesEditor.Entities
{
    public class RigaSpeseSkippata
    {
        public string Foglio { get; private set; }
        public int Riga { get; private set; }
        public int Colonna { get; private set; }
        public object DatoNonValido { get; private set; }

        public RigaSpeseSkippata(string foglio, int riga, int colonna, object datoErrato)
        {
            if (string.IsNullOrWhiteSpace(foglio))
                throw new ArgumentNullException(nameof(foglio));

            Foglio = foglio;
            Riga = riga;
            Colonna = colonna;
            DatoNonValido = datoErrato;
        }
    }
}