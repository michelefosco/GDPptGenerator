using ReportRefresher.Constants;
using ReportRefresher.Entities;
using ReportRefresher.Enums;
using System.Collections.Generic;
using System;
using ReportRefresher.Helpers;
using System.Linq;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// Produzione contenuti extra per il debug
    /// </summary>
    internal class Step_ProduzioneContenutiExtraPerFileDebug: Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            GeneraFormulePer_ReportisticaPerCategoria(context.Configurazione, context.DebugInfoLogger, context.CategorieFornitori, context.FornitoriCensitiInReport);
            context.DebugInfoLogger.LogText("Generazione formule per copia e incolla in 'Reportistica per categoria'", "OK");

            return null;
        }

        private void GeneraFormulePer_ReportisticaPerCategoria(Configurazione configurazione, FileDebugHelper debugInfoLogger, List<string> categorieFornitori, List<FornitoreCensito> fornitoriCensiti)
        {
            debugInfoLogger.LogFormuleReportisticaPerCategoriaIntestazione();
            foreach (var categoriaFornitori in categorieFornitori)
            {
                // Colonna J
                GeneraFormulaPer_ReportisticaPerCategoria_Categoria(
                    debugInfoLogger: debugInfoLogger,
                    fornitoriCensiti: fornitoriCensiti,
                    categoria: categoriaFornitori,
                    colonnaFormula: configurazione.ReportisticaPerTipologia_Colonna_ConsumiSpeseAdOre_Ore,
                    testoDaConcatenareNellaFormula: configurazione.ReportisticaPerTipologia_Formula_ConsumiSpeseAdOre_Ore);

                // Colonna K
                GeneraFormulaPer_ReportisticaPerCategoria_Categoria(
                    debugInfoLogger: debugInfoLogger,
                    fornitoriCensiti: fornitoriCensiti,
                    categoria: categoriaFornitori,
                    colonnaFormula: configurazione.ReportisticaPerTipologia_Colonna_ConsumiSpeseAdOre_Euro,
                    testoDaConcatenareNellaFormula: configurazione.ReportisticaPerTipologia_Formula_ConsumiSpeseAdOre_Euro);

                // Colonna L
                GeneraFormulaPer_ReportisticaPerCategoria_Categoria(
                    debugInfoLogger: debugInfoLogger,
                    fornitoriCensiti: fornitoriCensiti,
                    categoria: categoriaFornitori,
                    colonnaFormula: configurazione.ReportisticaPerTipologia_Colonna_ConsumiSpeseLumpSum_Euro,
                    testoDaConcatenareNellaFormula: configurazione.ReportisticaPerTipologia_Formula_ConsumiSpeseLumpSum_Euro);

                // Colonna N
                GeneraFormulaPer_ReportisticaPerCategoria_Categoria(
                    debugInfoLogger: debugInfoLogger,
                    fornitoriCensiti: fornitoriCensiti,
                    categoria: categoriaFornitori,
                    colonnaFormula: configurazione.ReportisticaPerTipologia_Colonna_AllocateSpeseAdOre_Ore,
                    testoDaConcatenareNellaFormula: configurazione.ReportisticaPerTipologia_Formula_AllocateSpeseAdOre_Ore);

                // Colonna O
                GeneraFormulaPer_ReportisticaPerCategoria_Categoria(
                    debugInfoLogger: debugInfoLogger,
                    fornitoriCensiti: fornitoriCensiti,
                    categoria: categoriaFornitori,
                    colonnaFormula: configurazione.ReportisticaPerTipologia_Colonna_AllocateSpeseAdOre_Euro,
                    testoDaConcatenareNellaFormula: configurazione.ReportisticaPerTipologia_Formula_AllocateSpeseAdOre_Euro);

                // Colonna P
                GeneraFormulaPer_ReportisticaPerCategoria_Categoria(
                    debugInfoLogger: debugInfoLogger,
                    fornitoriCensiti: fornitoriCensiti,
                    categoria: categoriaFornitori,
                    colonnaFormula: configurazione.ReportisticaPerTipologia_Colonna_AllocateSpeseLumpSum_Euro,
                    testoDaConcatenareNellaFormula: configurazione.ReportisticaPerTipologia_Formula_AllocateSpeseLumpSum_Euro);
            }
        }
        private void GeneraFormulaPer_ReportisticaPerCategoria_Categoria(FileDebugHelper debugInfoLogger, List<FornitoreCensito> fornitoriCensiti, string categoria, int colonnaFormula, string testoDaConcatenareNellaFormula)
        {
            var porzioneDiFormula = 1;
            var testoformula = "'=0";

            // scorre tutte le sigle dei fornitori
            foreach (var siglaFornitore in fornitoriCensiti.Where(_ => _.Categoria.Equals(categoria, StringComparison.OrdinalIgnoreCase)).Select(_ => _.SiglaInReport).OrderBy(s => s))
            {
                var nuovoPezzoDaConcatenare = string.Format(testoDaConcatenareNellaFormula, siglaFornitore);
                if (testoformula.Length + nuovoPezzoDaConcatenare.Length > Numbers.LIMITE_LUNGHEZZA_FORMULE_EXCEL)
                {
                    // ho raggiunto il limite della lunghezza della formula consentita da Excel:
                    // scrivo la formula calcolata finora e creo una nuova forzione di formula
                    debugInfoLogger.LogFormuleReportisticaPerCategoriaRiga(categoria, ((ColumnIDS)colonnaFormula).GetEnumDescription(), testoformula, porzioneDiFormula);

                    testoformula = "'=0";
                    porzioneDiFormula++;
                }
                testoformula += nuovoPezzoDaConcatenare;
            }

            debugInfoLogger.LogFormuleReportisticaPerCategoriaRiga(categoria, ((ColumnIDS)colonnaFormula).GetEnumDescription(), testoformula, porzioneDiFormula);
        }
        //private void ChiusuraFileDebug(FileDebugHelper debugInfoLogger, UpdateReportsOutput updateReportsOutput)
        //{
        //    if (debugInfoLogger != null)
        //    {
        //        debugInfoLogger.LogUpdateReportsOutput(updateReportsOutput);
        //        debugInfoLogger.Beautify();
        //        debugInfoLogger.Save();
        //    }
        //}
    }
}