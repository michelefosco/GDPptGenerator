using System;

namespace ReportRefresher.Entities
{
    internal class RigaLogElaborazioniTabellaSintesi
    {
        public readonly string Reparto;
        public readonly string Operazione;
        public readonly double Valore;
        public readonly string SimboloValore;
        public readonly int Mese;
        public readonly string Fornitore;
        public readonly int Riga;

        public RigaLogElaborazioniTabellaSintesi(string reparto, string operazione, double valore, string simboloValore, int mese, string fornitore, int riga)
        {
            if (string.IsNullOrWhiteSpace(reparto)) { throw new ArgumentNullException(nameof(reparto)); }
            if (string.IsNullOrWhiteSpace(operazione)) { throw new ArgumentNullException(nameof(operazione)); }
            if (valore == 0) { throw new ArgumentNullException(nameof(valore)); }
            if (string.IsNullOrWhiteSpace(simboloValore)) { throw new ArgumentNullException(nameof(simboloValore)); }
            if (mese < 1 || mese > 12) { throw new ArgumentNullException(nameof(mese)); }
            if (string.IsNullOrWhiteSpace(fornitore)) { throw new ArgumentNullException(nameof(fornitore)); }
            if (riga < 1) { throw new ArgumentNullException(nameof(riga)); }

            Reparto = reparto;
            Operazione = operazione;
            Valore = valore;
            SimboloValore = simboloValore;
            Mese = mese;
            Fornitore = fornitore;
            Riga = riga;
        }
    }
}
