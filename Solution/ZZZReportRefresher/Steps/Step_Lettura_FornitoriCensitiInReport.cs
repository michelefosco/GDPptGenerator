using ReportRefresher.Entities;
using ReportRefresher.Entities.Exceptions;
using ReportRefresher.Enums;
using System.Collections.Generic;
using System;
using System.Linq;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Lettura info fornitori censiti
    /// </summary>
    internal class Step_Lettura_FornitoriCensitiInReport : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            // Lettura Sigla e Nome da "Anagrafica fornitori" dal file "Report"
            context.FornitoriCensitiInReport = GetListaFornitoriCensiti(context.InfoFileReport, context.Configurazione);
            context.UpdateReportsOutput.SettaFornitoriCensiti(context.FornitoriCensitiInReport);

            return null;
        }

        private List<FornitoreCensito> GetListaFornitoriCensiti(InfoFileReport infoFileReport, Configurazione configurazione)
        {
            var worksheetName = infoFileReport.WorksheetName_AnagraficaFornitori; // Anagrafica fornitori

            var _listaFornitori = new List<FornitoreCensito>();

            var ultimaRigaUsataNelFoglio = infoFileReport.EPPlusHelper.GetRowsLimit(worksheetName);
            var rigaCorrente = configurazione.AnagraficaFornitori_PrimaRigaFornitori;
            while (rigaCorrente <= ultimaRigaUsataNelFoglio)
            {
                var siglaFornitore = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.AnagraficaFornitori_ColonnaSigle);
                var nomeFornitore = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.AnagraficaFornitori_ColonnaNomi);

                // se tutti i valori sono nulli ho raggiunto la fine della tabella, mi fermo
                if (siglaFornitore == null && nomeFornitore == null)
                { break; }

                // se uno o più valori è nullo sollevo un'eccezione
                if (string.IsNullOrWhiteSpace(siglaFornitore) || string.IsNullOrWhiteSpace(nomeFornitore))
                {
                    //TEST:SituazioniNonValide.InputFile_Report_EUG.GD_INPUT_EXPORT_REPORT_KO_Nome_Fornitore_Mancante()
                    //  InputReport_KO_NomeFornitoreMancante.xlsx
                    //TEST: SituazioniNonValide.InputFile_Report_EUG.GD_INPUT_EXPORT_REPORT_KO_Sigla_Fornitore_Mancante() Line 90	C#
                    //  InputReport_KO_SiglaFornitoreMancante.xlsx
                    throw new ManagedException(
                        tipologiaErrore: TipologiaErrori.DatoMancante,
                        tipologiaCartella: TipologiaCartelle.ReportInput,
                        nomeDatoErrore: siglaFornitore == null ? NomiDatoErrore.SiglaFornitore : NomiDatoErrore.NomeFornitore,
                        worksheetName: worksheetName,
                        rigaCella: rigaCorrente,
                        colonnaCella: siglaFornitore == null ? configurazione.AnagraficaFornitori_ColonnaSigle : configurazione.AnagraficaFornitori_ColonnaNomi,
                        dato: null,
                        percorsoFile: null
                        );
                }

                //Sigla fornitore non univoca nel file report
                if (_listaFornitori.Any(_ => _.SiglaInReport.Equals(siglaFornitore, StringComparison.InvariantCultureIgnoreCase)))
                    //TEST: SituazioniNonValide.InputFile_Report.GD_INPUT_EXPORT_CONTROLLER_KO_Reparto_NonUnivoco()
                    //  InputReport_KO_SiglaFornitoreNonUnivoco.xlsx
                    //TEST: SituazioniNonValide.InputFile_Report_EUG.GD_INPUT_EXPORT_REPORT_KO_Sigla_Fornitore_NonUnivoco() Line 26	C#
                    //  InputReport_KO_SiglaFornitoreNonUnivoco.xlsx
                    throw new ManagedException(
                        tipologiaErrore: TipologiaErrori.DatoNonUnivoco,
                        tipologiaCartella: TipologiaCartelle.ReportInput,
                        nomeDatoErrore: NomiDatoErrore.SiglaFornitore,
                        worksheetName: worksheetName,
                        rigaCella: rigaCorrente,
                        colonnaCella: configurazione.AnagraficaFornitori_ColonnaSigle,
                        dato: siglaFornitore,
                        percorsoFile: null
                        );

                //Nome fornitore non univoco nel file report
                if (_listaFornitori.Any(_ => _.HasThisName(nomeFornitore)))
                    //TEST:SituazioniNonValide.InputFile_Report_EUG.GD_INPUT_EXPORT_REPORT_KO_Nome_Fornitore_NonUnivoco()
                    //InputReport_KO_NomeFornitoreNonUnivoco.xlsx
                    throw new ManagedException(
                        tipologiaErrore: TipologiaErrori.DatoNonUnivoco,
                        tipologiaCartella: TipologiaCartelle.ReportInput,
                        nomeDatoErrore: NomiDatoErrore.NomeFornitore,
                        worksheetName: worksheetName,
                        rigaCella: rigaCorrente,
                        colonnaCella: configurazione.AnagraficaFornitori_ColonnaNomi,
                        dato: nomeFornitore,
                        percorsoFile: null
                        );

                _listaFornitori.Add(new FornitoreCensito(siglaInReport: siglaFornitore,
                                                        nomeSuController: nomeFornitore,
                                                        presenteSoloInListaDati: false)
                                                        );

                // vado alla riga successiva
                rigaCorrente++;
            }

            return _listaFornitori;
        }

    }
}