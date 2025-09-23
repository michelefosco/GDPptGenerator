using ReportRefresher.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportRefresher.Entities
{
    public class FornitoreNonCensito
    {
        public FornitoreNonCensito(string nomeSuController)
        {
            if (string.IsNullOrWhiteSpace(nomeSuController))
            { throw new ArgumentNullException(nameof(nomeSuController)); }
            NomeSuController = nomeSuController;

            FornitoriNonCensiti_InfoFoglioReparto = new List<FornitoreNonCensito_InfoFoglioReparto>();
        }
        public string NomeSuController { get; private set; }

        public void AssociaFoglioAndRepartoAlNomeDelNuovoFornitore(string foglio, string reparto, TipologieDiSpesa tipologiaDiSpesa)
        {
            if (!FornitoriNonCensiti_InfoFoglioReparto.Any(fr => fr.Foglio.Equals(foglio, StringComparison.CurrentCultureIgnoreCase)
                                                        && fr.Reparto.Equals(reparto, StringComparison.CurrentCultureIgnoreCase)))
            {
                FornitoriNonCensiti_InfoFoglioReparto.Add(new FornitoreNonCensito_InfoFoglioReparto(foglio, reparto, tipologiaDiSpesa));
            }
        }
        public List<FornitoreNonCensito_InfoFoglioReparto> FornitoriNonCensiti_InfoFoglioReparto { get; private set; }
    }
}