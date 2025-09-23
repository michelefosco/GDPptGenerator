using ReportRefresher.Enums;
using System;

namespace ReportRefresher.Entities
{
    public class RigaSpese
    {
        public RigaSpese(string centroDiCosto, string nomeReparto, FornitoreCensito fornitore, TipologieDiSpesa tipologieDiSpesa, StatusSpesa statusSpesa, double spesa, double? ore, double? costoOrarioApplicato, DateTime dataInizio, DateTime dataFine)
        {
            if (string.IsNullOrWhiteSpace(centroDiCosto))
            { throw new ArgumentNullException(nameof(centroDiCosto)); }
            if (string.IsNullOrWhiteSpace(nomeReparto))
            { throw new ArgumentNullException(nameof(nomeReparto)); }
            if (fornitore == null)
            { throw new ArgumentNullException(nameof(fornitore)); }
            if (dataInizio > dataFine)
            { throw new ArgumentOutOfRangeException("Data inizio non può essere maggiore di data fine nell'intervallo di spese"); }

            CentroDiCosto = centroDiCosto;
            NomeReparto = nomeReparto;
            Fornitore = fornitore;
            TipologiaDiSpesa = tipologieDiSpesa;
            StatusSpesa = statusSpesa;
            Spesa = spesa;
            Ore = ore;
            CostoOrarioApplicato = costoOrarioApplicato;
            DataInizio = dataInizio;
            DataFine = dataFine;
            //
            _setFieldsForTabellaSintesi(out DataInizioTabellaSintesi, out DataFineTabellaSintesi, out NumeroMesiSplitSpesaInTabellaSintesi);
        }

        // Centro di costo a cui è inputata la spesa
        readonly public string CentroDiCosto;

        // Nome del reparto a cui è afferente il centro di costo
        readonly public string NomeReparto;

        // Fornitore a cui è stata pagata o verrà pagata la spesa
        readonly public FornitoreCensito Fornitore;

        // Tipologie Di Spesa ("ORE", "LUMP SUM")
        readonly public TipologieDiSpesa TipologiaDiSpesa;

        // StatusSpesa ("Actual", "Commitment")
        readonly public StatusSpesa StatusSpesa;

        // Importo speso
        readonly public double Spesa;

        // Numero di ore erogate dal fornitore
        readonly public double? Ore;

        // Per le spese di tipo ad Ore viene indicato anche il costo orario applicato alla spesa corrente
        readonly public double? CostoOrarioApplicato;

        // Data di inizio del periodo a cui si riferisce la spesa
        readonly public DateTime DataInizio;

        // Data di fine del periodo a cui si riferisce la spesa
        readonly public DateTime DataFine;

        #region Campi per i calcoli necessari nella tabella Sintesi
        readonly public DateTime DataInizioTabellaSintesi;
        readonly public DateTime DataFineTabellaSintesi;
        readonly public int NumeroMesiSplitSpesaInTabellaSintesi;
        private void _setFieldsForTabellaSintesi(out DateTime dataInizioTabellaSintesi, out DateTime dataFineTabellaSintesi, out int numeroMesiSplitSpesaInTabellaSintesi)
        {
            if (DataInizio.Year == DataFine.Year && DataInizio.Month == DataFine.Month)
            {
                // Per le date nello stesso mese l'intervallo diventa inizio e fine del mese
                dataInizioTabellaSintesi = new DateTime(DataInizio.Year, DataInizio.Month, 1); // primo giorno del mese;
                dataFineTabellaSintesi = new DateTime(DataFine.Year, DataFine.Month, 1).AddMonths(1).AddDays(-1); // ultimo giorno del mese                                                                                                                 
            }
            else
            {
                // Le date sono su due mesi diversi
                dataInizioTabellaSintesi = (DataInizio.Day <= 15)
                        ? new DateTime(DataInizio.Year, DataInizio.Month, 1) // primo giorno del mese
                        : new DateTime(DataInizio.Year, DataInizio.Month, 1).AddMonths(1); // primo giorno del mese successivo
                dataFineTabellaSintesi = (DataFine.Day <= 15)
                        ? new DateTime(DataFine.Year, DataFine.Month, 1).AddDays(-1) // ultimo giorno del mese precedente
                        : new DateTime(DataFine.Year, DataFine.Month, 1).AddMonths(1).AddDays(-1); // ultimo giorno del mese

                if (dataInizioTabellaSintesi >= dataFineTabellaSintesi)
                {
                    // le date determinate con le regole standard NON sono valide
                    // rielaboro sulla base del mese con più giorni nel periodo
                    var numeroGiorniNelMeseDiInizio = (new DateTime(DataInizio.Year, DataInizio.Month, 1).AddMonths(1) - DataInizio).TotalDays;
                    var numeroGiorniNelMeseDiFine = DataFine.Day;

                    dataInizioTabellaSintesi = (numeroGiorniNelMeseDiInizio > numeroGiorniNelMeseDiFine)
                                ? new DateTime(DataInizio.Year, DataInizio.Month, 1) // primo giorno del mese
                                : new DateTime(DataInizio.Year, DataInizio.Month, 1).AddMonths(1); // primo giorno del mese successivo
                    dataFineTabellaSintesi = (numeroGiorniNelMeseDiInizio > numeroGiorniNelMeseDiFine)
                                ? new DateTime(DataFine.Year, DataFine.Month, 1).AddDays(-1) // ultimo giorno del mese precedente
                                : new DateTime(DataFine.Year, DataFine.Month, 1).AddMonths(1).AddDays(-1); // ultimo giorno del mese
                }
            }

            // Calcolo il numero di mesi nell'intervallo
            //  A) se entambe le date sono nello stesso anno allora è sufficiente guardare la differenza tra i numeri di mese
            //  B) in questo caso, sommiamo il numero di mesi per arrivare alla fine dell'anno per la data di inizio e i mesi trascorsi nell'anno successivo per la data di fine
            numeroMesiSplitSpesaInTabellaSintesi = (dataInizioTabellaSintesi.Year == dataFineTabellaSintesi.Year)
                    ? dataFineTabellaSintesi.Month - dataInizioTabellaSintesi.Month + 1 // caso A
                    : 12 - dataInizioTabellaSintesi.Month + 1 + dataFineTabellaSintesi.Month; // caso B
        }
        #endregion
    }
}