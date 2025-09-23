using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportRefresher.Entities
{
    public class FornitoreCensito
    {

        private Dictionary<string, double> _costiOrarioPerReparto;

        // un fornitore con "_hasMultipleMultiName" a true è un fornitore il cui nome è cambiato nel tempo, quindi
        // per praticità nella tabella "Anagrafica fornitore" il suo vecchio e nuovo nome convivono separati da un ';'
        // es. "LEFO s.r.l.;LEFO 2 s.r.l."
        private readonly bool _hasMultipleMultiName;

        /// <summary>
        /// Sigla del fornitore che lo identifica all'interno del file report
        /// </summary>
        public string SiglaInReport { get; private set; }

        /// <summary>
        /// Nome del fornitore che lo identifica all'interno file controller
        /// </summary>
        public string NomeSuController { get; private set; }
        public string Categoria { get; private set; }

        /// <summary>
        /// Costo orario da applicarsi per tutti i reparti che non hanno un prezzo specifico
        /// </summary>
        public double CostoOrarioStandard { get; private set; }
        public bool HasCostoOrarioSettato { get; private set; }

        /// <summary>
        /// Alcuni fornitori vengono inseriti per comodità solo in "Lista dati"
        /// </summary>
        public bool PresenteSoloInListaDati { get; private set; }

        /// <summary>
        /// Indica che il fornitore deve essere presente nei report in quanto ad esso sono associate delle spese
        /// </summary>
        public bool DeveEsserePresenteNeiReport { get; private set; }

        public FornitoreCensito(string siglaInReport, string nomeSuController, bool presenteSoloInListaDati = false)
        {
            if (string.IsNullOrWhiteSpace(siglaInReport))
                throw new ArgumentNullException(nameof(siglaInReport));
            if (string.IsNullOrWhiteSpace(nomeSuController))
                throw new ArgumentNullException(nameof(nomeSuController));

            SiglaInReport = siglaInReport;
            NomeSuController = nomeSuController;
            _hasMultipleMultiName = nomeSuController.Contains(';');
            PresenteSoloInListaDati = presenteSoloInListaDati;
            DeveEsserePresenteNeiReport = !presenteSoloInListaDati;

            if (presenteSoloInListaDati && SiglaInReport != NomeSuController)
            {
                throw new ArgumentException("I fornitori indicati come 'PresenteSoloInListaDati' devo avere necessariamente SiglaInReport e NomeSuController uguali");
            }
        }

        public bool HasThisName(string name)
        {
            if (_hasMultipleMultiName)
            {
                return NomeSuController.ToUpper().Contains(name.ToUpper());
            }
            else
            {
                return NomeSuController.Equals(name, StringComparison.OrdinalIgnoreCase);
            }
        }

        public void SettaCategoria(string categoria)
        {
            Categoria = categoria;
        }

        public void SettaCostiOrari(double costoOrarioStandard, Dictionary<string, double> costiOrarioPerReparto)
        {
            // Francesco dice: Il costo orario potrebbe essere a zero in quanto potrebbero
            // inserire il valore zero in caso in cui il fornitore lavori in modalità Lump Sum
            if (costoOrarioStandard < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(costoOrarioStandard) + " i costi orari devono essere maggiori di zero");
            }

            // Il costo orario per reparto non può essere a zero, se indicato deve essere positivo
            if (costiOrarioPerReparto != null && costiOrarioPerReparto.Any(_ => _.Value <= 0))
            {
                throw new ArgumentOutOfRangeException(nameof(costiOrarioPerReparto) + " i costi orari devono essere maggiori di zero");
            }

            HasCostoOrarioSettato = true;
            CostoOrarioStandard = costoOrarioStandard;
            _costiOrarioPerReparto = costiOrarioPerReparto;
        }

        public void Setta_DeveEsserePresenteNeiReport()
        {
            DeveEsserePresenteNeiReport = true;
        }

        public bool HasRepartCostoOrarioCustom(string nomeReparto)
        {
            ThrowExceptionIf_CostoOrario_NotSet();

            if (string.IsNullOrWhiteSpace(nomeReparto))
                throw new ArgumentNullException(nameof(nomeReparto));

            if (_costiOrarioPerReparto == null)
                return false;

            return _costiOrarioPerReparto.ContainsKey(nomeReparto);
        }

        public double GetCostoOrarioPerReparto(string nomeReparto)
        {
            ThrowExceptionIf_CostoOrario_NotSet();

            if (string.IsNullOrWhiteSpace(nomeReparto))
                throw new ArgumentNullException(nameof(nomeReparto));

            if (_costiOrarioPerReparto == null)
                return CostoOrarioStandard;

            if (_costiOrarioPerReparto.ContainsKey(nomeReparto))
                return _costiOrarioPerReparto[nomeReparto];

            return CostoOrarioStandard;
        }

        public List<string> GetRepartiConCostiCustom()
        {
            ThrowExceptionIf_CostoOrario_NotSet();

            if (_costiOrarioPerReparto != null && _costiOrarioPerReparto.Any())
            {
                return _costiOrarioPerReparto.Keys.ToList();
            }
            return null;
        }

        private void ThrowExceptionIf_CostoOrario_NotSet()
        {
            if (!HasCostoOrarioSettato)
            {
                throw new Exception("Costi orario non ancora settati. Utilizzare prima il metodo SettaCostiOrari");
            }
        }
    }
}