using System.ComponentModel;

namespace ReportRefresher.Enums
{
    public enum NomiDatoErrore
    {
        None = 0,//Nome dato non impostato

        [Description("Sigla fornitore")]
        SiglaFornitore,

        [Description("Nome fornitore")]
        NomeFornitore,

        [Description("Categoria fornitore")]
        CategoriaFornitore,

        [Description("Costo orario fornitore")]
        CostoOrarioFornitore,

        [Description("Nome reparto")]
        NomeReparto,

        [Description("Centro di costo")]
        CentroDiCosto,

        [Description("Periodo")]
        Periodo,

        [Description("Spesa")]
        Spesa,

        [Description("Tipologia spesa")]
        TipologiaSpesa,

        [Description("Intestazione dati")]
        IntestazioneDati,

        [Description("Formula")]
        Formula,

        [Description("Tupla")]
        Tupla,

        [Description("Data inizio periodo spesa")]
        DataInizioPeriodoSpesa,

        [Description("Data fine periodo spesa")]
        DataFinePeriodoSpesa
    }
}