using System;

namespace ReportRefresher.Entities
{
    public class RigaTabellaConsumiSpacchettati
    {
        public RigaTabellaConsumiSpacchettati(string reparto,
                                double spesaAdOre_Actual_Ore,
                                double spesaAdOre_Actual_Euro,
                                double spesaAdOre_Commitment_Euro,
                                double spesaAdOre_Commitment_Ore,
                                double spesaLumpSum_Actual,
                                double spesaLumpSum_Commitment)
        {
            if (string.IsNullOrWhiteSpace(reparto))
            { throw new ArgumentNullException(nameof(reparto)); }

            Reparto = reparto;
            SpesaAdOre_Actual_Ore = spesaAdOre_Actual_Ore;
            SpesaAdOre_Actual_Euro = spesaAdOre_Actual_Euro;
            SpesaAdOre_Commitment_Ore = spesaAdOre_Commitment_Ore;
            SpesaAdOre_Commitment_Euro = spesaAdOre_Commitment_Euro;
            SpesaLumpSum_Actual = spesaLumpSum_Actual;
            SpesaLumpSum_Commitment = spesaLumpSum_Commitment;
        }

        public readonly string Reparto;
        public readonly double SpesaAdOre_Actual_Ore;
        public readonly double SpesaAdOre_Actual_Euro;
        public readonly double SpesaAdOre_Commitment_Ore;
        public readonly double SpesaAdOre_Commitment_Euro;
        public readonly double SpesaLumpSum_Actual;
        public readonly double SpesaLumpSum_Commitment;

        public bool HasSpese
        {
            get
            {
                return (SpesaAdOre_Actual_Euro != 0 || SpesaAdOre_Commitment_Euro != 0 || SpesaLumpSum_Commitment != 0);
            }
        }
    }
}