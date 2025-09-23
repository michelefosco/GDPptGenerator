using EPPlusExtensions;
using ReportRefresher.Constants;
using ReportRefresher.Entities;
using ReportRefresher.Entities.Exceptions;
using ReportRefresher.Enums;
using System.IO;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Verifiche preliminare sui file di input "Controller"
    /// </summary>
    internal class Step_Start_InfoFileController : Step_Start_InfoFile_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            context.InfoFileController = BuildInfoFileController(context.UpdateReportsInput.FileController_FilePath);
            context.DebugInfoLogger.LogText("Verifiche sul file 'Controller'", "OK");

            return null;
        }

        private InfoFileController BuildInfoFileController(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ManagedException(
                tipologiaErrore: TipologiaErrori.DatoMancante,
                tipologiaCartella: TipologiaCartelle.Controller,
                messaggioPerUtente: MessaggiErrorePerUtente.PercorsoFileMancante,
                percorsoFile: filePath
                );
            }

            // Verifica info su "Export Controller" file: percorso esistente
            if (!File.Exists(filePath))
            {
                //TEST: SituazioniNonValide.InputFile_Controller.PercorsoFile_Inisistente()
                //C:\NomeInesistente.xlsx
                throw new ManagedException(
                tipologiaErrore: TipologiaErrori.FileMancante,
                tipologiaCartella: TipologiaCartelle.Controller,
                messaggioPerUtente: MessaggiErrorePerUtente.FileMancante,
                percorsoFile: filePath
                );
            }

            var epPlusHelper = new EPPlusHelper();
            // Verifica info su "Export Controller" file: si apre correttamente
            if (!epPlusHelper.Open(filePath))
            {
                //TEST: SituazioniNonValide.InputFile_Controller.FileInFormatoNonCorretto()
                //FoglioExcelInFormatoSbagliato.xlsx
                throw new ManagedException(
                    tipologiaErrore: TipologiaErrori.FormatoFileErrato,
                    tipologiaCartella: TipologiaCartelle.Controller,
                    messaggioPerUtente: MessaggiErrorePerUtente.FileImpossibileDaAprire,
                    percorsoFile: filePath
                    );
            }

            #region Verifica info su "Export Controller" file: continene tutti i worksheet necessari
            var worksheetNamesList = epPlusHelper.GetWorksheetNames();
            var worksheetName_Recap = GetNomeWorhSheetNellaCartella(epPlusHelper,TipologiaCartelle.Controller, worksheetNamesList, "RECAP", "", "");
            var worksheetName_ActualSoloCdc = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.Controller, worksheetNamesList, "ACT SOLO CDC", "", "");
            // questo foglio ha un nome variabile, ma deve iniziare sempre con "FBL3N_act"
            var worksheetName_FBL3Nact = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.Controller, worksheetNamesList, "", "FBL3N_ACT", "");
            var worksheetName_CommitmentSoloCdc = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.Controller, worksheetNamesList, "COMMITMENT SOLO CDC", "", "");
            var worksheetName_ME5A = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.Controller, worksheetNamesList, "ME5A", "", "");
            var worksheetNam_ME2N = GetNomeWorhSheetNellaCartella(epPlusHelper, TipologiaCartelle.Controller, worksheetNamesList, "ME2N", "", "");
            #endregion

            return new InfoFileController(epPlusHelper, filePath, worksheetName_Recap, worksheetName_ActualSoloCdc, worksheetName_FBL3Nact, worksheetName_CommitmentSoloCdc, worksheetName_ME5A, worksheetNam_ME2N);
        }
    }
}