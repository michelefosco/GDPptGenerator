using ReportRefresher.Constants;
using ReportRefresher.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportRefresher.Helpers
{
    public class ValueHelper
    {
        public static bool CastNomeFornitoreDaColonnaConDoppioValore(object objectFornitore, out string nome)
        {
            var fornitoreFormatoStringa = objectFornitore.ToString().Trim();
            var posizionePrimoBlank = fornitoreFormatoStringa.IndexOf(' ');
            if (posizionePrimoBlank < 1)
            {
                nome = null;
                return false;
            }

            // Prende la parte di stringa dal primo blanck in poi. Questa parte rappresenta la parte
            // di nome, a cui vengono comunque poi tolti gli eventuali blank all'inizio e alla fine.
            nome = fornitoreFormatoStringa.Substring(posizionePrimoBlank).Trim();

            return true;
        }

        public static bool CastTipologieDiSpesa(object objectTipologieDiSpesa, out TipologieDiSpesa? tipologieDiSpesa)
        {
            if (objectTipologieDiSpesa != null)
            {
                var textVersione = objectTipologieDiSpesa.ToString().Trim().ToUpper();

                // verifico che la Tipologia di spesa abbia un valore valido (ovvero "ORE" o "LUMP SUM")
                if (textVersione.Equals(TipologieDiSpesa.LumpSum.GetEnumDescription()))
                {
                    tipologieDiSpesa = TipologieDiSpesa.LumpSum;
                    return true;
                }

                if (textVersione.Equals(TipologieDiSpesa.AdOre.GetEnumDescription()))
                {
                    tipologieDiSpesa = TipologieDiSpesa.AdOre;
                    return true;
                }
            }

            // non ho trovato nessuna corrispondenza
            tipologieDiSpesa = null;
            return false;
        }



        /// <summary>
        /// Metodo per aggiungere delta positivi ad una lista di valori.
        /// I "Valori da modificare" sono solo quelli maggiori di zero, in assenza di quest'ultimi si usano tutti.
        /// Il delta viene equidistribuitio tra i "Valori da modificare"
        /// Questa è la versione richiesta da GD
        /// </summary>
        public static List<double> AddDeltaToValues_PriorityForNumberGreaterThanZero(List<double> inNumbers, double delta)
        {
            if (Math.Round(delta, Numbers.NumeroDecimaliImportiSpese) < 0)
            { throw new ArgumentOutOfRangeException(nameof(delta)); }

            var outNumbers = NumberContainer.BuildListFromDoubles(inNumbers);
            
            // prendo solo i numeri maggiori di zero, in loro assenza li prendo tutti
            var valuesToBeIncreased = outNumbers.Any(_ => _.Value > 0)
                            ? outNumbers.Where(_ => _.Value > 0)
                            : outNumbers;

            var valueToBeAdded = delta / valuesToBeIncreased.Count();

            foreach (var valueToBeIncreased in valuesToBeIncreased)
            {
                valueToBeIncreased.Value += valueToBeAdded;
            }

            return outNumbers.Select(_ => Math.Round(_.Value, Numbers.NumeroDecimaliImportiSpese)).ToList();
        }

        public static List<double> AddDeltaToValues_ConPareggio(List<double> inNumbers, double delta)
        {
            if (Math.Round(delta, Numbers.NumeroDecimaliImportiSpese) < 0)
            { throw new ArgumentOutOfRangeException(nameof(delta)); }

            var outNumbers = NumberContainer.BuildListFromDoubles(inNumbers);

            // 
            while (Math.Round(delta, Numbers.NumeroDecimaliImportiSpese) > 0)
            {
                var minimo = outNumbers.Min(_ => _.Value);
                var valoriMinimi = outNumbers.Where(_ => _.Value == minimo).ToList();

                foreach (var value in valoriMinimi)
                {
                    var valoriMaggioriDelMinimo = outNumbers.Where(_ => _.Value > minimo).ToList();

                    double valoreDaAggiungere;
                    if (valoriMaggioriDelMinimo.Any())
                    {
                        // abbiamo tra le righe da modificare un valore più alto del minimo
                        // aggiungo quindi alle righe con il valore minimo il delta per raggiungere questo secondo valore
                        var valoreSecondoMinimo = valoriMaggioriDelMinimo.Min(_ => _.Value);

                        var valoreDistribuitoTraValoriDaModificare = delta / valoriMinimi.Count();

                        var differenzaTraMimino_e_SecondoMinimo = valoreSecondoMinimo - minimo;

                        valoreDaAggiungere = (valoreDistribuitoTraValoriDaModificare < differenzaTraMimino_e_SecondoMinimo)
                                            ? valoreDistribuitoTraValoriDaModificare
                                            : differenzaTraMimino_e_SecondoMinimo;
                    }
                    else
                    {
                        // il valore attuale delle righe è uguale per tutti (tutte uguali al minimo)
                        // posso equidistribuire il delta tra le celle
                        valoreDaAggiungere = delta / valoriMinimi.Count;
                    }

                    value.Value += valoreDaAggiungere;

                    // aggiorno il valore del delta
                    delta -= valoreDaAggiungere;
                }
            }

            return outNumbers.Select(_ => Math.Round(_.Value, Numbers.NumeroDecimaliImportiSpese)).ToList();
        }


        public static List<double> RemoveDeltaFromValues(List<double> inNumbers, double delta, bool negativeAllowed, out double remainingDelta)
        {
            if (Math.Round(delta, Numbers.NumeroDecimaliImportiSpese) < 0)
            { throw new ArgumentOutOfRangeException(nameof(delta)); }

            var outNumbers = NumberContainer.BuildListFromDoubles(inNumbers);

            while (Math.Round(delta, Numbers.NumeroDecimaliImportiSpese) > 0)
            {
                // seleziono solo le righe con un valore positivo
                var righeModificabili = outNumbers.Any(_ => _.Value > 0)
                                    ? outNumbers.Where(_ => _.Value > 0).ToList()
                                    : outNumbers.ToList();

                var valoreDaSottrarreInModoEquidistribuito = delta / righeModificabili.Count;
                var valoreMinimoDisponibile = righeModificabili.Min(_ => _.Value);

                if (valoreMinimoDisponibile == 0 && !negativeAllowed)
                {
                    break;
                }

                var valoreDaTogliere = (valoreDaSottrarreInModoEquidistribuito <= valoreMinimoDisponibile || valoreMinimoDisponibile == 0)
                                    ? valoreDaSottrarreInModoEquidistribuito
                                    : valoreMinimoDisponibile;

                foreach (var riga in righeModificabili)
                {
                    riga.Value -= valoreDaTogliere;

                    // aggiorno il valore del delta
                    delta -= valoreDaTogliere;
                }
            }

            remainingDelta = delta;
            return outNumbers.Select(_ => Math.Round(_.Value, Numbers.NumeroDecimaliImportiSpese)).ToList();
        }

        private class NumberContainer
        {
            public NumberContainer(double value)
            {
                Value = value;
            }

            public double Value;
            public static List<NumberContainer> BuildListFromDoubles(List<double> inNumbers)
            {
                var outNumbers = new List<NumberContainer>();
                foreach (var value in inNumbers)
                {
                    outNumbers.Add(new NumberContainer(value));
                }
                return outNumbers;
            }
        }
    }
}