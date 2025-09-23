using ReportRefresher.Enums;
using System;

namespace ReportRefresher.Entities
{
    public class TuplaSpesePerAnno
    {
        public TuplaSpesePerAnno(string siglaFornitore, string reparto, TipologieDiSpesa tipologiaDiSpesa)
        {
            if (string.IsNullOrWhiteSpace(siglaFornitore))
                throw new ArgumentNullException(nameof(siglaFornitore));
            if (string.IsNullOrWhiteSpace(reparto))
                throw new ArgumentNullException(nameof(reparto));


            Reparto = reparto;
            SiglaFornitore = siglaFornitore;
            TipologiaDiSpesa = tipologiaDiSpesa;
            //
            SpeseActualPerMese = new double[12];
            SpeseCommittedPerMese = new double[12];
        }

        // Chiave composita Reparto+SiglaFornitore+TipologiaDiSpesa
        public readonly string SiglaFornitore;
        public readonly string Reparto;
        public readonly TipologieDiSpesa TipologiaDiSpesa;
        //
        public double[] SpeseActualPerMese;
        public double[] SpeseCommittedPerMese;
    }
}