using System;
using System.Collections.Generic;

namespace ReportRefresher.Entities
{
    public class Reparto
    {
        public Reparto() { }

        public Reparto(string nome)
        {
            if (string.IsNullOrWhiteSpace(nome))
                throw new ArgumentNullException(nameof(nome));

            Nome = nome;
            CentriDiCosto = new List<string>();
        }

        public void AggiungiCentriDiCosto(string centroDiCosto)
        {
            CentriDiCosto.Add(centroDiCosto);
        }

        public void AggiungiCentriDiCosto(List<string> centriDiCosto)
        {
            CentriDiCosto.AddRange(centriDiCosto);
        }

        /// <summary>
        /// Utilizzato in caso di merge
        /// </summary>
        public void AzzeraCentriDiCosto()
        {
            CentriDiCosto.Clear();
        }

        public string Nome { get; private set; }
        public List<string> CentriDiCosto { get; private set; }
    }
}