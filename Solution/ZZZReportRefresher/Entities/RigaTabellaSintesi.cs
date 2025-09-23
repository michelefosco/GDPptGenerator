using ReportRefresher.Enums;
using System;

namespace ReportRefresher.Entities
{
    public class RigaTabellaSintesi
    {
        public RigaTabellaSintesi(string siglaFornitore, string reparto, TipologieDiSpesa tipologiaDiSpesa, string categoriaFornitori)
        {
            if (string.IsNullOrWhiteSpace(siglaFornitore))
                throw new ArgumentNullException(nameof(siglaFornitore));
            if (string.IsNullOrWhiteSpace(reparto))
                throw new ArgumentNullException(nameof(reparto));
            if (string.IsNullOrWhiteSpace(categoriaFornitori))
                throw new ArgumentNullException(nameof(categoriaFornitori));

            Reparto = reparto;
            SiglaFornitore = siglaFornitore;
            TipologiaDiSpesa = tipologiaDiSpesa;
            CategoriaFornitori = categoriaFornitori;
            //
            BudgetStanziato = new double[12];
            ValorePianificazioneConConsumi = new double[12];
            DeltaApplicato = new double[12];
            //
            AlreadyChecked = false;
        }

        public readonly string SiglaFornitore;
        public readonly string Reparto;
        public readonly TipologieDiSpesa TipologiaDiSpesa;
        public readonly string CategoriaFornitori;
        //
        public double[] BudgetStanziato;
        public double[] ValorePianificazioneConConsumi;
        public double[] DeltaApplicato;
        public double ExtraBudget;
        //
        public bool AlreadyChecked;
    }
}