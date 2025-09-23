using ReportRefresher.Enums;
using System;

namespace ReportRefresher.Entities
{
    public class FornitoreNonCensito_InfoFoglioReparto
    {
        public string Foglio { get; private set; }
        public string Reparto { get; private set; }
        public TipologieDiSpesa TipologiaDiSpesa { get; private set; }



        public FornitoreNonCensito_InfoFoglioReparto(string foglio, string reparto, TipologieDiSpesa tipologiaDiSpesa)
        {
            if (string.IsNullOrWhiteSpace(foglio))
                throw new ArgumentNullException(nameof(foglio));
            if (string.IsNullOrWhiteSpace(reparto))
                throw new ArgumentNullException(nameof(reparto));

            Foglio = foglio;
            Reparto = reparto;
            TipologiaDiSpesa = tipologiaDiSpesa;
        }
    }
}