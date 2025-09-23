using EPPlusExtensions;
using ReportRefresher.Constants;
using ReportRefresher.Entities;
using ReportRefresher.Entities.Exceptions;
using ReportRefresher.Enums;
using System.IO;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Verifiche preliminare sui file di input "Report"
    /// </summary>
    internal class Step_Start_InfoFileReport : Step_Start_InfoFile_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            context.InfoFileReport = BuildInfoFileReport(context.UpdateReportsInput.FileReport_FilePath);
            context.DebugInfoLogger.LogText("Verifiche sul file 'Report'", "OK");

            return null;
        }

        private InfoFileReport BuildInfoFileReport(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ManagedException(
                tipologiaErrore: TipologiaErrori.DatoMancante,
                tipologiaCartella: TipologiaCartelle.ReportInput,
                messaggioPerUtente: MessaggiErrorePerUtente.PercorsoFileMancante,
                percorsoFile: filePath
                );
            }

            // Verifica info su "Report" file: percorso esistente
            if (!File.Exists(filePath))
            {
                //TEST: SituazioniNonValide.InputFile_Report.PercorsoFile_Inisistente()
                //C:\\NomeInesistente.xlsx
                throw new ManagedException(
                tipologiaErrore: TipologiaErrori.FileMancante,
                tipologiaCartella: TipologiaCartelle.ReportInput,
                messaggioPerUtente: MessaggiErrorePerUtente.FileMancante,
                percorsoFile: filePath
                );
            }

            var epPlusHelper = new EPPlusHelper();
            // Verifica info su "Report" file: si apre correttamente
            if (!epPlusHelper.Open(filePath))
            {
                //TEST: SituazioniNonValide.InputFile_Report.FileInFormatoNonCorretto()
                //FoglioExcelInFormatoSbagliato.xlsx
                throw new ManagedException(
                    tipologiaErrore: TipologiaErrori.FormatoFileErrato,
                    tipologiaCartella: TipologiaCartelle.ReportInput,
                    messaggioPerUtente: MessaggiErrorePerUtente.FileImpossibileDaAprire,
                    percorsoFile: filePath
                    );
            }

            #region Verifica info su "Report" file: continene tutti i worksheet necessari
            var worksheetNamesList = epPlusHelper.GetWorksheetNames();
            var worksheetName_Avanzamento = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.ReportInput, worksheetNamesList, "AVANZAMENTO", "", "");
            var worksheetName_ConsumiSpacchettati = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.ReportInput, worksheetNamesList, "CONSUMI SPACCHETTATI", "", "");
            var worksheetName_ListaDati = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.ReportInput, worksheetNamesList, "LISTA DATI", "", "");
            var worksheetName_AnagraficaFornitori = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.ReportInput, worksheetNamesList, "ANAGRAFICA FORNITORI", "", "");
            // IL foglio "202x Sintesi" ha un nome variabile, ma finisce sempre con "sintesi"
            var worksheetName_Sintesi_EndsWith = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.ReportInput, worksheetNamesList, "", "", "SINTESI");
            var worksheetName_PianificazioneConConsumi = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.ReportInput, worksheetNamesList, "PIANIFICAZIONE CON CONSUMI", "", "");
            var worksheetName_SintesiLogModifiche = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.ReportInput, worksheetNamesList, "SINTESI-LOG MODIFICHE", "", "");
            var worksheetName_SpeseActual = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.ReportInput, worksheetNamesList, "SPESE ACTUAL", "", "");
            var worksheetName_SpeseComittment = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.ReportInput, worksheetNamesList, "SPESE COMMITMENT", "", "");
            var worksheetName_SpeseTotali = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.ReportInput, worksheetNamesList, "SPESE TOTALI", "", "");
            var worksheetName_ReportisticaPerTipologia = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.ReportInput, worksheetNamesList, "REPORTISTICA PER TIPOLOGIA", "", "");
            var worksheetName_BudgetStudiIpotesi = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.ReportInput, worksheetNamesList, "BUDGET STUDI -  IPOTESI 1", "", "");
            var worksheetName_Reportistica = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.ReportInput, worksheetNamesList, "REPORTISTICA", "", "");
            var worksheetName_PrevisioneAfinire = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.ReportInput, worksheetNamesList, "PREVISIONE A FINIRE", "", "");
            #endregion

            return new InfoFileReport(
                    epPlusHelper: epPlusHelper,
                    filePath: filePath,
                    worksheetName_Avanzamento: worksheetName_Avanzamento,
                    worksheetName_ConsumiSpacchettati: worksheetName_ConsumiSpacchettati,
                    worksheetName_ListaDati: worksheetName_ListaDati,
                    worksheetName_AnagraficaFornitori: worksheetName_AnagraficaFornitori,
                    worksheetName_Sintesi: worksheetName_Sintesi_EndsWith,
                    worksheetName_PianificazioneConConsumi: worksheetName_PianificazioneConConsumi,
                    worksheetName_SintesiLogModifiche: worksheetName_SintesiLogModifiche,
                    worksheetName_ReportisticaPerTipologia: worksheetName_ReportisticaPerTipologia,
                    worksheetName_SpeseActual: worksheetName_SpeseActual,
                    worksheetName_SpeseComittment: worksheetName_SpeseComittment,
                    worksheetName_SpeseTotali: worksheetName_SpeseTotali,
                    worksheetName_BudgetStudiIpotesi: worksheetName_BudgetStudiIpotesi,
                    worksheetName_Reportistica: worksheetName_Reportistica,
                    worksheetName_PrevisioneAfinire: worksheetName_PrevisioneAfinire);
        }
    }
}