using ReportRefresher.Entities;
using ReportRefresher.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Gestione tabella "Avanzamento"
    /// </summary>
    internal class Step_TabellaAvanzamento : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            context.RigheTabellaAvanzamento = CalcolaDatiTabellaAvanzamento(context.RepartiCensitiInReport, context.FornitoriCensitiInReport, context.RigheSpese);
            context.UpdateReportsOutput.SettaTabellaAvanzamento(context.RigheTabellaAvanzamento);
            context.DebugInfoLogger.LogRigheTabellaAvanzamento(context.RigheTabellaAvanzamento);
            context.DebugInfoLogger.LogText("Calcolo righe tabella Avanzamento", context.RigheTabellaAvanzamento.Count);

            AggiornaTabellaAvanzamento(context.InfoFileReport, context.Configurazione, context.RigheTabellaAvanzamento, context.UpdateReportsInput.DataAggiornamento);
            context.DebugInfoLogger.LogText("Aggiornamento tabella Avanzamento", "OK");

            return null;
        }

        private  List<RigaTabellaAvanzamento> CalcolaDatiTabellaAvanzamento(List<Reparto> repartiCensitiInReport, List<FornitoreCensito> fornitoriCensiti, List<RigaSpese> righeSpese)
        {
            var righeTabellaAvanzamento = new List<RigaTabellaAvanzamento>();

            // Vengono prodotte tutte le combinazioni di reparto e fornitore
            foreach (var nomeReparto in repartiCensitiInReport.Select(r => r.Nome).OrderBy(nome => nome))
            {
                // Dei fornitori devono 
                foreach (var fornitore in fornitoriCensiti.Where(_ => _.DeveEsserePresenteNeiReport).OrderBy(f => f.NomeSuController))
                {
                    // Spese "Ad Ore" - Somma della Spesa
                    var totaleSpeseAdOreEURO = righeSpese.Where(_ => _.NomeReparto == nomeReparto && _.Fornitore.SiglaInReport == fornitore.SiglaInReport && _.TipologiaDiSpesa == TipologieDiSpesa.AdOre)
                                        .Sum(_ => _.Spesa);
                    // Spese "Ad Ore" - Somma delle Ore
                    var totaleSpesaAdOreORE = righeSpese.Where(_ => _.NomeReparto == nomeReparto && _.Fornitore.SiglaInReport == fornitore.SiglaInReport && _.TipologiaDiSpesa == TipologieDiSpesa.AdOre)
                                        .Sum(_ => _.Ore) ?? 0;
                    // Spese "Lump sum" - Somma della Spesa
                    var totaleSpeseLumpSum = righeSpese.Where(_ => _.NomeReparto == nomeReparto && _.Fornitore.SiglaInReport == fornitore.SiglaInReport && _.TipologiaDiSpesa == TipologieDiSpesa.LumpSum)
                                        .Sum(_ => _.Spesa);

                    var rigaAvanzamento = new RigaTabellaAvanzamento(
                        nomeReparto,                // Reparto
                        fornitore.NomeSuController, // Fornitore.Nome
                        fornitore.SiglaInReport,    // Fornitore.Sigla
                        totaleSpeseAdOreEURO,       // Euro Act + Commit
                        totaleSpesaAdOreORE,        // ORE Act + Commit
                        totaleSpeseLumpSum          // EURO LUMP Act + Commit
                        );
                    righeTabellaAvanzamento.Add(rigaAvanzamento);
                }
            }
            return righeTabellaAvanzamento;
        }
        private  void AggiornaTabellaAvanzamento(InfoFileReport infoFileReport, Configurazione configurazione, List<RigaTabellaAvanzamento> righeTabellaAvanzamento, DateTime dataAggiornamento)
        {
            var worksheetName = infoFileReport.WorksheetName_Avanzamento; // "Avanzamento"

            // Svuotamento delle Colonne E-J nella tabella "Avanzamento"
            var ultimaRigaUsataNelFoglio = infoFileReport.EPPlusHelper.GetRowsLimit(worksheetName);
            infoFileReport.EPPlusHelper.CleanCellsContent(worksheetName,
                                    configurazione.Avanzamento_PrimaRigaDaAggiornare, configurazione.Avanzamento_PrimaColonnaDaAggiornare,   //Cella in alto a sinistra della selezione da cancellare //E2
                                    ultimaRigaUsataNelFoglio, configurazione.Avanzamento_UltimaColonnaDaAggiornare);     //Cella in basso a destra  J<x>            


            // Scrittura delle righe nella tabella "AVANZAMENTO"
            var currentRow = configurazione.Avanzamento_PrimaRigaDaAggiornare;
            foreach (var rigaTabella in righeTabellaAvanzamento)
            {
                infoFileReport.EPPlusHelper.SetValuesOnRow(worksheetName, currentRow, configurazione.Avanzamento_PrimaColonnaDaAggiornare,
                            rigaTabella.Reparto,                //#1 Reparto
                            rigaTabella.NomeFornitore,          //#2 Nome fornitore
                            rigaTabella.SiglaFornitore,         //#3 Sigla fornitore
                            rigaTabella.TotaleSpeseAdOre_Euro,  //#4 Euro Act + Commit
                            rigaTabella.TotaleSpeseAdOre_Ore,   //#5 ORE Act + Commit    
                            rigaTabella.TotaleSpeseLumpSum      //#6 EURO LUMP Act + Commit
                            );
                currentRow++;
            }

            // Setto la data nella cella M1 con la data dell'ultimo aggiornamento
            infoFileReport.EPPlusHelper.SetValue(worksheetName, configurazione.Avanzamento_AddressCellaUltimoAggiornamento, dataAggiornamento);
        }
    }
}