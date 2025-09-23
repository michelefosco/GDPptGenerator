using System;

namespace ReportRefresher.Entities
{
    public class RigaTabellaAvanzamento
    {
        public RigaTabellaAvanzamento(string reparto,
                string nomeFornitore,
                string siglaFornitore,
                double totaleSpeseAdOre_Euro,
                double totaleSpeseAdOre_Ore,
                double totaleSpeseLumpSum
            )
        {
            if (string.IsNullOrWhiteSpace(reparto))
                throw new ArgumentNullException(nameof(reparto));
            if (string.IsNullOrWhiteSpace(nomeFornitore))
                throw new ArgumentNullException(nameof(nomeFornitore));
            if (string.IsNullOrWhiteSpace(siglaFornitore))
                throw new ArgumentNullException(nameof(siglaFornitore));

            Reparto = reparto;
            NomeFornitore = nomeFornitore;
            SiglaFornitore = siglaFornitore;
            TotaleSpeseAdOre_Euro = totaleSpeseAdOre_Euro;
            TotaleSpeseAdOre_Ore = totaleSpeseAdOre_Ore;
            TotaleSpeseLumpSum = totaleSpeseLumpSum;
        }

        public readonly string Reparto;
        public readonly string NomeFornitore;
        public readonly string SiglaFornitore;
        public readonly double TotaleSpeseAdOre_Euro; //Euro Act+Commit 
        public readonly double TotaleSpeseAdOre_Ore;  //Ore Act+Commit
        public readonly double TotaleSpeseLumpSum;    //Euro LUMP Act+Commit
    }
}