using ReportRefresher.Constants;
using ReportRefresher.Entities.Exceptions;
using ReportRefresher.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Classe base per le verifiche preliminari sui file di input "Controller" e "Report"
    /// </summary>
    internal abstract class Step_Start_InfoFile_Base : Step_Base
    {
        internal string GetNomeWorhSheetNellaCartella(EPPlusExtensions.EPPlusHelper epPlusHelper, TipologiaCartelle tipologiaCartella, List<string> worksheetNames, string nomeEsatto, string startWith = null, string endsWith = null)
        {
            if (string.IsNullOrEmpty(nomeEsatto) && string.IsNullOrEmpty(startWith) && string.IsNullOrEmpty(endsWith))
            {
                throw new ArgumentException("E necessario speficare almento un argomento di ricerca");
            }

            var nomeFoglioTrovato = string.Empty;

            if (!string.IsNullOrEmpty(nomeEsatto))
            {
                // ricerca per nome esatto
                if (!worksheetNames.Contains(nomeEsatto))
                {
                    throw new ManagedException(
                            tipologiaErrore: TipologiaErrori.FoglioMancante,
                            tipologiaCartella: tipologiaCartella,
                            messaggioPerUtente: string.Format(MessaggiErrorePerUtente.FoglioMancante, tipologiaCartella.ToString(), nomeEsatto),
                            worksheetName: nomeEsatto);
                }
                nomeFoglioTrovato = nomeEsatto;
            }
            if (!string.IsNullOrEmpty(startWith))
            {
                // ricerca per nome che cominicia con...
                nomeFoglioTrovato = worksheetNames.FirstOrDefault(_ => _.StartsWith(startWith));
                if (string.IsNullOrEmpty(nomeFoglioTrovato))
                {
                    //TEST: SituazioniNonValide.InputFile_Controller.FileConWorksheetMancante(string inputExportControllerfileName)
                    //InputExportController_KO_NomeMancante_2.xlsx
                    throw new ManagedException(
                            tipologiaErrore: TipologiaErrori.FoglioMancante,
                            tipologiaCartella: tipologiaCartella,
                            messaggioPerUtente: string.Format(MessaggiErrorePerUtente.FoglioMancante, tipologiaCartella.ToString(), startWith + "?????"),
                            worksheetName: startWith);
                }
            }
            if (!string.IsNullOrEmpty(endsWith))
            {
                // ricerca per nome che finisce con...
                nomeFoglioTrovato = worksheetNames.FirstOrDefault(_ => _.Trim().EndsWith(endsWith));
                if (string.IsNullOrEmpty(nomeFoglioTrovato))
                {
                    //TEST: SituazioniNonValide.InputFile_Report.FileConWorksheetMancante(string inputExportControllerfileName)
                    //InputReport_KO_NomeMancante_7.xlsx
                    throw new ManagedException(
                            tipologiaErrore: TipologiaErrori.FoglioMancante,
                            tipologiaCartella: tipologiaCartella,
                            messaggioPerUtente: string.Format(MessaggiErrorePerUtente.FoglioMancante, tipologiaCartella.ToString(), "?????" + endsWith),
                            worksheetName: endsWith);
                }
            }

            // todo: check if the worksheet has at least one cell

            if (!epPlusHelper.HasWorksheetCells(nomeFoglioTrovato))
            {
                throw new ManagedException(
                    tipologiaErrore: TipologiaErrori.FoglioIncompleto,
                    tipologiaCartella: tipologiaCartella,
                    messaggioPerUtente: string.Format(MessaggiErrorePerUtente.FoglioIncompleto, tipologiaCartella.ToString(), nomeEsatto),
                    worksheetName: nomeFoglioTrovato);
            }

            return nomeFoglioTrovato;
        }
    }
}