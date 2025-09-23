using System;

namespace ReportRefresher.Entities
{
    public class RigaSpeseSkippata
    {
        public string Foglio { get; private set; }
        public int Riga { get; private set; }
        public int Colonna { get; private set; }
        public object DatoNonValido { get; private set; }
        public string Reparto { get; private set; }

        public RigaSpeseSkippata(string foglio, int riga, int colonna, string reparto, object datoErrato)
        {
            if (string.IsNullOrWhiteSpace(foglio))
                throw new ArgumentNullException(nameof(foglio));

            if (string.IsNullOrWhiteSpace(reparto))
                throw new ArgumentNullException(nameof(reparto));


            Foglio = foglio;
            Riga = riga;
            Colonna = colonna;
            DatoNonValido = datoErrato;
            Reparto = reparto;
        }
    }
}