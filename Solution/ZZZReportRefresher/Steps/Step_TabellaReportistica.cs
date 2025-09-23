using ReportRefresher.Entities;
using ReportRefresher.Entities.Exceptions;
using ReportRefresher.Enums;
using System.Collections.Generic;
using System;
using System.Globalization;
using System.Linq;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Gestione tabella "Reportistica"
    /// </summary>
    internal class Step_TabellaReportistica : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            AggiornaFormuleSuTabellaReportistica(context.InfoFileReport, context.Configurazione, context.RepartiCensitiInReport, context.RigheTabellaConsumiSpacchettati);
            NascondiRigheInEccessoSuTabellaReportistica(context.InfoFileReport, context.Configurazione);
            context.DebugInfoLogger.LogText("Aggiornamento delle formule sul foglio reportistica", "OK");

            return null;
        }

        private  void AggiornaFormuleSuTabellaReportistica(InfoFileReport infoFileReport, Configurazione configurazione, List<Reparto> repartiCensitiInReport, List<RigaTabellaConsumiSpacchettati> consumiSpacchettati)
        {
            var worksheetName = infoFileReport.WorksheetName_Reportistica;  // Reportistica

            #region Verifico i nomi presenti nella colonna dei reparti, mi assicuro che siano tutti validi
            var repartiTrovatiSulFoglioReportistica = new List<string>();
            var rigaCorrente = configurazione.Reportistica_PrimaRigaResponsabili;
            while (true)
            {
                var nomeReparto = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.Reportistica_ColonnaResponsabili);
                if (string.IsNullOrWhiteSpace(nomeReparto) || nomeReparto.Equals("Totale", StringComparison.CurrentCultureIgnoreCase))   // mi interrompo con l'ultima cella con valore nullo
                { break; }

                // verifico che il reparto si tra quelli censiti
                var repartoCensito = repartiCensitiInReport.SingleOrDefault(_ => _.Nome.Equals(nomeReparto, StringComparison.CurrentCultureIgnoreCase));
                if (repartoCensito == null)
                {
                    throw new ManagedException(
                        tipologiaErrore: TipologiaErrori.DatoNonValido,
                        tipologiaCartella: TipologiaCartelle.ReportInput,
                        messaggioPerUtente: $"Il foglio '{worksheetName}' ha un reparto non correttamente censito nel foglio {infoFileReport.WorksheetName_ListaDati}')",
                        worksheetName: worksheetName,
                        rigaCella: rigaCorrente,
                        colonnaCella: configurazione.Reportistica_ColonnaResponsabili,
                        dato: nomeReparto,
                        nomeDatoErrore: NomiDatoErrore.NomeReparto,
                        percorsoFile: null
                        );
                }

                repartiTrovatiSulFoglioReportistica.Add(repartoCensito.Nome);
                rigaCorrente++;
            }
            #endregion

            #region Verifico che il foglio "Reportistica" abbia tutti i reparti previsti
            // A questo punto ho verificato che tutti i reparti trovati in "Reportistica" sono presenti tra quelli censiti.
            // Verifico quindi che tutti i reparti censiti (in "Lista dati") e che abbiano delle spese inputate a loro
            // siano effettivamente presenti in "Reportistica"
            foreach (var repartoCensitoInReport in repartiCensitiInReport)
            {
                if (!repartiTrovatiSulFoglioReportistica.Any(_ => _.Equals(repartoCensitoInReport.Nome, StringComparison.CurrentCultureIgnoreCase)))
                {
                    if (consumiSpacchettati.Single(_ => _.Reparto.Equals(repartoCensitoInReport.Nome, StringComparison.CurrentCultureIgnoreCase)).HasSpese)
                    {
                        // abbiamo un reparto censito in "Lista dati", non presenti in "Reportistica", ma con delle spese
                        // in questo caso il reparto dovrebbe essere nella lista di "Reportistica" ma risulta mancante.
                        throw new ManagedException(
                            tipologiaErrore: TipologiaErrori.DatoMancante,
                            tipologiaCartella: TipologiaCartelle.ReportInput,
                            messaggioPerUtente: $"Nel foglio '{worksheetName}' risulta mancare il reparto '{repartoCensitoInReport.Nome}')",
                            worksheetName: worksheetName,
                            rigaCella: null,
                            colonnaCella: null,
                            dato: repartoCensitoInReport.Nome,
                            nomeDatoErrore: NomiDatoErrore.NomeReparto,
                            percorsoFile: null
                            );
                    }
                }
            }
            #endregion

            #region Incollo la formula con segnaposto nella prima riga della tabella
            int primaRigaConDati = configurazione.Reportistica_PrimaRigaResponsabili;
            // Setto la formula per la colonna "Resto ore"
            infoFileReport.EPPlusHelper.SetFormula(worksheetName, primaRigaConDati, configurazione.Reportistica_ColonnaRestoOre, configurazione.Reportistica_FormulaPerColonnaRestoOre);
            // Setto la formula per la colonna "Resto euro"
            infoFileReport.EPPlusHelper.SetFormula(worksheetName, primaRigaConDati, configurazione.Reportistica_ColonnaRestoEuro, configurazione.Reportistica_FormulaPerColonnaRestoEuro);
            // Setto la formula per la colonna "Resto lump"
            infoFileReport.EPPlusHelper.SetFormula(worksheetName, primaRigaConDati, configurazione.Reportistica_ColonnaRestoLump, configurazione.Reportistica_FormulaPerColonnaRestoLump);
            #endregion

            #region Copio e incollo la formula con segnaposto nel resto delle righe, in questo modo la formula verrà aggiornata con i rifermienti correnti 
            // per incollare parto dalla seconda riga nella tabella
            rigaCorrente = configurazione.Reportistica_PrimaRigaResponsabili + 1;
            while (true)
            {
                var nomeReparto = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.Reportistica_ColonnaResponsabili);
                if (string.IsNullOrWhiteSpace(nomeReparto) || nomeReparto.Equals("Totale", StringComparison.CurrentCultureIgnoreCase))   // mi interrompo con l'ultima cella con valore nullo
                { break; }

                // Incollo la formula per la colonna "Resto ore"
                infoFileReport.EPPlusHelper.CopiaFormula(worksheetName, primaRigaConDati, configurazione.Reportistica_ColonnaRestoOre,
                                                                    rigaCorrente, configurazione.Reportistica_ColonnaRestoOre);

                // Incollo la formula per la colonna "Resto euro"
                infoFileReport.EPPlusHelper.CopiaFormula(worksheetName, primaRigaConDati, configurazione.Reportistica_ColonnaRestoEuro,
                                                                    rigaCorrente, configurazione.Reportistica_ColonnaRestoEuro);

                // Incollo la formula per la colonna "Resto lump"
                infoFileReport.EPPlusHelper.CopiaFormula(worksheetName, primaRigaConDati, configurazione.Reportistica_ColonnaRestoLump,
                                                                    rigaCorrente, configurazione.Reportistica_ColonnaRestoLump);

                rigaCorrente++;
            }
            #endregion

            #region Sostituzione del valore segnaposto con il valore calcolato
            rigaCorrente = configurazione.Reportistica_PrimaRigaResponsabili;
            // i numeri nelle formule devono essere in formato US, come le formule devono essere in inglese
            var nfi = new CultureInfo("en-US", false).NumberFormat;
            while (true)
            {
                var nomeReparto = infoFileReport.EPPlusHelper.GetString(worksheetName, rigaCorrente, configurazione.Reportistica_ColonnaResponsabili);
                if (string.IsNullOrWhiteSpace(nomeReparto) || nomeReparto.Equals("Totale", StringComparison.CurrentCultureIgnoreCase))   // mi interrompo con l'ultima cella con valore nullo
                { break; }

                var rigaConsumiSpacchettati = consumiSpacchettati.Single(_ => _.Reparto.Equals(nomeReparto, StringComparison.CurrentCultureIgnoreCase));

                // Aggiorno la formula per la colonna "Resto ore"
                var valorePerFormula = (rigaConsumiSpacchettati.SpesaAdOre_Actual_Ore + rigaConsumiSpacchettati.SpesaAdOre_Commitment_Ore).ToString("0.00", nfi);
                var currentFormula = infoFileReport.EPPlusHelper.GetFormula(worksheetName, rigaCorrente, configurazione.Reportistica_ColonnaRestoOre);
                var newFormula = currentFormula.Replace(configurazione.Reportistica_ValoreSegnapostoPerFormule, valorePerFormula);
                infoFileReport.EPPlusHelper.SetFormula(worksheetName, rigaCorrente, configurazione.Reportistica_ColonnaRestoOre, newFormula);

                // Aggiorno la formula per la colonna "Resto euro"
                valorePerFormula = (rigaConsumiSpacchettati.SpesaAdOre_Actual_Euro + rigaConsumiSpacchettati.SpesaAdOre_Commitment_Euro).ToString("0.00", nfi);
                currentFormula = infoFileReport.EPPlusHelper.GetFormula(worksheetName, rigaCorrente, configurazione.Reportistica_ColonnaRestoEuro);
                newFormula = currentFormula.Replace(configurazione.Reportistica_ValoreSegnapostoPerFormule, valorePerFormula);
                infoFileReport.EPPlusHelper.SetFormula(worksheetName, rigaCorrente, configurazione.Reportistica_ColonnaRestoEuro, newFormula);

                // Aggiorno la formula per la colonna "Resto lump"                
                valorePerFormula = (rigaConsumiSpacchettati.SpesaLumpSum_Actual + rigaConsumiSpacchettati.SpesaLumpSum_Commitment).ToString("0.00", nfi);
                currentFormula = infoFileReport.EPPlusHelper.GetFormula(worksheetName, rigaCorrente, configurazione.Reportistica_ColonnaRestoLump);
                newFormula = currentFormula.Replace(configurazione.Reportistica_ValoreSegnapostoPerFormule, valorePerFormula);
                infoFileReport.EPPlusHelper.SetFormula(worksheetName, rigaCorrente, configurazione.Reportistica_ColonnaRestoLump, newFormula);
                rigaCorrente++;
            }
            #endregion
        }

        private  void NascondiRigheInEccessoSuTabellaReportistica(InfoFileReport infoFileReport, Configurazione configurazione)
        {
            var worksheetName = infoFileReport.WorksheetName_Reportistica;  // Reportistica

            var primaRigaDaNascondere = infoFileReport.EPPlusHelper.GetFirstEmptyRow(
                worksheetName: worksheetName,
                rowToStartSearchFrom: configurazione.Reportistica_PrimaRigaNascondibile,
                colToBeChecked: configurazione.Reportistica_ColonnaSigle_SX
                );

            var ultimaDaNascondere = infoFileReport.EPPlusHelper.GetFirstRowWithAnyValue(
                    worksheetName: worksheetName,
                    rowToStartSearchFrom: primaRigaDaNascondere,
                    colToBeChecked: configurazione.Reportistica_ColonnaSigle_SX + 1, // la prima colonna a destra delle sigle
                    ignoreThisText: "non-visibile"
                    );

            if (ultimaDaNascondere > 0)
            {
                // Mostro quelle che dovrebbero essere visibili
                infoFileReport.EPPlusHelper.NascondiRighe(worksheetName, configurazione.Reportistica_PrimaRigaNascondibile, primaRigaDaNascondere, false);
                // Nascondo quelle che dovrebbero essere nascoste
                infoFileReport.EPPlusHelper.NascondiRighe(worksheetName, primaRigaDaNascondere, ultimaDaNascondere, true);
            }
        }
    }
}