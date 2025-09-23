using EPPlusExtensions;
using System;
using System.Drawing;

namespace ReportRefresher.Entities
{
    public class InfoFileReport : InfoFileBase
    {
        public readonly string WorksheetName_ListaDati;
        public readonly string WorksheetName_Avanzamento;
        public readonly string WorksheetName_ConsumiSpacchettati;
        public readonly string WorksheetName_AnagraficaFornitori;
        public readonly string WorksheetName_Sintesi;
        public readonly string WorksheetName_PianificazioneConConsumi;
        public readonly string WorksheetName_SintesiLogModifiche;
        public readonly string WorksheetName_SpeseActual;
        public readonly string WorksheetName_SpeseComittment;
        public readonly string WorksheetName_SpeseTotali;
        public readonly string WorksheetName_ReportisticaPerTipologia;
        public readonly string WorksheetName_BudgetStudiIpotesi;
        public readonly string WorksheetName_Reportistica;
        public readonly string WorksheetName_PrevisioneAfinire;
        public InfoFileReport(EPPlusHelper epPlusHelper,
                        string filePath,
                        string worksheetName_Avanzamento,
                        string worksheetName_ConsumiSpacchettati,
                        string worksheetName_ListaDati,
                        string worksheetName_AnagraficaFornitori,
                        string worksheetName_Sintesi,
                        string worksheetName_PianificazioneConConsumi,
                        string worksheetName_SintesiLogModifiche,
                        string worksheetName_SpeseActual,
                        string worksheetName_SpeseComittment,
                        string worksheetName_SpeseTotali,
                        string worksheetName_ReportisticaPerTipologia,
                        string worksheetName_BudgetStudiIpotesi,
                        string worksheetName_Reportistica,
                        string worksheetName_PrevisioneAfinire) : base(epPlusHelper, filePath)
        {
            if (string.IsNullOrWhiteSpace(worksheetName_Avanzamento)) throw new ArgumentNullException(nameof(worksheetName_Avanzamento));
            if (string.IsNullOrWhiteSpace(worksheetName_ConsumiSpacchettati)) throw new ArgumentNullException(nameof(worksheetName_ConsumiSpacchettati));
            if (string.IsNullOrWhiteSpace(worksheetName_ListaDati)) throw new ArgumentNullException(nameof(worksheetName_ListaDati));
            if (string.IsNullOrWhiteSpace(worksheetName_AnagraficaFornitori)) throw new ArgumentNullException(nameof(worksheetName_AnagraficaFornitori));
            if (string.IsNullOrWhiteSpace(worksheetName_Sintesi)) throw new ArgumentNullException(nameof(worksheetName_Sintesi));
            if (string.IsNullOrWhiteSpace(worksheetName_PianificazioneConConsumi)) throw new ArgumentNullException(nameof(worksheetName_PianificazioneConConsumi));
            if (string.IsNullOrWhiteSpace(worksheetName_SintesiLogModifiche)) throw new ArgumentNullException(nameof(worksheetName_SintesiLogModifiche));
            if (string.IsNullOrWhiteSpace(worksheetName_SpeseActual)) throw new ArgumentNullException(nameof(worksheetName_SpeseActual));
            if (string.IsNullOrWhiteSpace(worksheetName_SpeseComittment)) throw new ArgumentNullException(nameof(worksheetName_SpeseComittment));
            if (string.IsNullOrWhiteSpace(worksheetName_SpeseTotali)) throw new ArgumentNullException(nameof(worksheetName_SpeseTotali));
            if (string.IsNullOrWhiteSpace(worksheetName_ReportisticaPerTipologia)) throw new ArgumentNullException(nameof(worksheetName_ReportisticaPerTipologia));
            if (string.IsNullOrWhiteSpace(worksheetName_BudgetStudiIpotesi)) throw new ArgumentNullException(nameof(worksheetName_BudgetStudiIpotesi));
            if (string.IsNullOrWhiteSpace(worksheetName_Reportistica)) throw new ArgumentNullException(nameof(worksheetName_Reportistica));
            if (string.IsNullOrWhiteSpace(worksheetName_PrevisioneAfinire)) throw new ArgumentNullException(nameof(worksheetName_PrevisioneAfinire));

            WorksheetName_Avanzamento = worksheetName_Avanzamento;
            WorksheetName_ConsumiSpacchettati = worksheetName_ConsumiSpacchettati;
            WorksheetName_ListaDati = worksheetName_ListaDati;
            WorksheetName_AnagraficaFornitori = worksheetName_AnagraficaFornitori;
            WorksheetName_Sintesi = worksheetName_Sintesi;
            WorksheetName_PianificazioneConConsumi = worksheetName_PianificazioneConConsumi;
            WorksheetName_SintesiLogModifiche = worksheetName_SintesiLogModifiche;
            WorksheetName_SpeseActual = worksheetName_SpeseActual;
            WorksheetName_SpeseComittment = worksheetName_SpeseComittment;
            WorksheetName_SpeseTotali = worksheetName_SpeseTotali;
            WorksheetName_ReportisticaPerTipologia = worksheetName_ReportisticaPerTipologia;
            WorksheetName_BudgetStudiIpotesi = worksheetName_BudgetStudiIpotesi;
            WorksheetName_Reportistica = worksheetName_Reportistica;
            WorksheetName_PrevisioneAfinire = worksheetName_PrevisioneAfinire;
        }
    }
}