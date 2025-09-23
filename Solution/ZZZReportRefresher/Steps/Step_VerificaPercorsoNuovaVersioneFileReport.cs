using ReportRefresher.Constants;
using ReportRefresher.Entities;
using ReportRefresher.Entities.Exceptions;
using ReportRefresher.Enums;
using System.IO;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Salvataggio del file report nella sua nuova versione
    /// </summary>
    internal class Step_VerificaPercorsoNuovaVersioneFileReport : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            // Verifica che il file necessario come output non sia già esistente
            if (File.Exists(context.UpdateReportsInput.NewReport_FilePath))
            {
                throw new ManagedException(
                    tipologiaErrore: TipologiaErrori.FileGiaEsistente,
                    tipologiaCartella: TipologiaCartelle.ReportOutput,
                    messaggioPerUtente: MessaggiErrorePerUtente.FileGiaEsistente,
                    percorsoFile: context.UpdateReportsInput.NewReport_FilePath);
            }

            return null;
        }
    }
}