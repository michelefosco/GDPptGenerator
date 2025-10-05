using FilesEditor.Entities;
using FilesEditor.Enums;
namespace FilesEditor.Helpers
{
    public class ConfigurazioneHelper
    {
        public static Configurazione GetConfigurazioneDefault()
        {
            var configurazione = new Configurazione();

            // Filtri
            configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_FILTERS_FIRST_ROW = 4;
            configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_FILTERS_TABLE_COL = (int)ColumnIDS.K;
            configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_FILTERS_FIELD_COL = (int)ColumnIDS.L;

            // Slide da generare
            configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_FIRST_ROW = 4;
            configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_POWERPOINTFILE_COL = (int)ColumnIDS.A;
            configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_TITLE_COL = (int)ColumnIDS.B;
            configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_CONTENT_1_COL = (int)ColumnIDS.C;
            configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_CONTENT_2_COL = (int)ColumnIDS.D;
            configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_CONTENT_3_COL = (int)ColumnIDS.E;
            configurazione.DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_LAYOUT_COL = (int)ColumnIDS.F;

            //#region Cartella "Controller"
            //// Foglio "ACT SOLO CDC"
            //configurazione.ActualSoloCdc_IntestazionePrimaColonna = "Chiave po";
            //configurazione.ActualSoloCdc_ColonnaCentriDiCosto = (int)ColumnIDS.E;
            //configurazione.ActualSoloCdc_ColonnaSpesa = (int)ColumnIDS.G;
            //configurazione.ActualSoloCdc_ColonnaFornitore = (int)ColumnIDS.R;
            //configurazione.ActualSoloCdc_ColonnaDataInizio = (int)ColumnIDS.S;
            //configurazione.ActualSoloCdc_ColonnaDataFine = (int)ColumnIDS.U;

            //// Foglio "commitment solo cdc"
            //configurazione.CommitmentSoloCdc_IntestazionePrimaColonna = "Chiave po";
            //configurazione.CommitmentSoloCdc_ColonnaCentriDiCosto = (int)ColumnIDS.E;
            //configurazione.CommitmentSoloCdc_ColonnaSpesa = (int)ColumnIDS.G;
            //configurazione.CommitmentSoloCdc_ColonnaFornitore = (int)ColumnIDS.P;
            //configurazione.CommitmentSoloCdc_ColonnaDataInizio = (int)ColumnIDS.Q;
            //configurazione.CommitmentSoloCdc_ColonnaDataFine = (int)ColumnIDS.S;


            //// Foglio "FBL3N_act"
            //configurazione.FBL3Nact_IntestazionePrimaColonna = "Chiave po";
            //configurazione.FBL3Nact_ColonnaSpesa = (int)ColumnIDS.L;
            //configurazione.FBL3Nact_ColonnaCentriDiCosto = (int)ColumnIDS.R;
            //configurazione.FBL3Nact_ColonnaFornitore = (int)ColumnIDS.T;
            //configurazione.FBL3Nact_ColonnaDataInizio = (int)ColumnIDS.U;
            //configurazione.FBL3Nact_ColonnaDataFine = (int)ColumnIDS.W;
            ////configurazione.FBL3Nact_IntestazionePrimaColonna = "G/L Account";


            //// Foglio "ME5A"
            //configurazione.ME5A_IntestazionePrimaColonna = "Chiave po";
            //configurazione.ME5A_ColonnaFornitore = (int)ColumnIDS.H;
            //configurazione.ME5A_ColonnaSpesa = (int)ColumnIDS.Q;
            //configurazione.ME5A_ColonnaCentriDiCosto = (int)ColumnIDS.R;
            //configurazione.ME5A_ColonnaDataInizio = (int)ColumnIDS.T;
            //configurazione.ME5A_ColonnaDataFine = (int)ColumnIDS.V;
            ////configurazione.ME5A_IntestazionePrimaColonna = "Purchase Requisition";


            //// Foglio "ME2N"
            //configurazione.ME2N_IntestazionePrimaColonna = "Chiave po";
            //configurazione.ME2N_ColonnaSpesa = (int)ColumnIDS.J;
            //configurazione.ME2N_ColonnaCentriDiCosto = (int)ColumnIDS.O;
            //configurazione.ME2N_ColonnaFornitore = (int)ColumnIDS.Q;
            //configurazione.ME2N_ColonnaDataInizio = (int)ColumnIDS.R;
            //configurazione.ME2N_ColonnaDataFine = (int)ColumnIDS.T;
            ////configurazione.ME2N_IntestazionePrimaColonna = "Purchasing Document";


            //// Foglio "Recap"
            //configurazione.Recap_ColonnaReparti = (int)ColumnIDS.B;
            //configurazione.Recap_ColonnaCentriDiCosto = (int)ColumnIDS.C;
            //configurazione.Recap_Riga_PrimaConDati = 5;
            //// I seguenti reparti: Milandri, Testoni, Gianese, Moretti devono fare tutti capo al reparto Lanzarini. Cioè per intenderci: dove trovi uno di quei nomi, puoi censirli tutti come se fossero Lanzarini. È un eccezione unica.
            //configurazione.Recap_ListaMergeReparti = new List<string> { "Milandri F.|Lanzarini L.", "Testoni L|Lanzarini L.", "Gianese G.|Lanzarini L.", "Moretti C.|Lanzarini L." };
            //#endregion

            //#region Cartella "Report"
            //configurazione.AttivazioneOpzione_RefreshOnLoad_SuTutteLePivotTable = true;

            //// Foglio "Anagrafica Fornitori"
            //configurazione.AnagraficaFornitori_ColonnaSigle = (int)ColumnIDS.A;
            //configurazione.AnagraficaFornitori_ColonnaNomi = (int)ColumnIDS.B;
            //configurazione.AnagraficaFornitori_PrimaRigaFornitori = 2;

            //// Foglio "Avanzamento"
            //configurazione.Avanzamento_PrimaRigaDaAggiornare = 2;
            //configurazione.Avanzamento_PrimaColonnaDaAggiornare = (int)ColumnIDS.E;
            //configurazione.Avanzamento_UltimaColonnaDaAggiornare = (int)ColumnIDS.J;
            //configurazione.Avanzamento_AddressCellaUltimoAggiornamento = "M1";

            //// Foglio "Consumi Spacchettati"
            //configurazione.ConsumiSpacchettati_ColonnaReparti = (int)ColumnIDS.E;
            //configurazione.ConsumiSpacchettati_ColonnaOreActual = (int)ColumnIDS.H;            
            //configurazione.ConsumiSpacchettati_PrimaRigaConDati = 16;

            //// Foglio "Ipotesi studio"
            //configurazione.IpotesiStudio_ColonnaSigle = (int)ColumnIDS.B;
            //configurazione.IpotesiStudio_ColonnaUltimaDaOrdinare = (int)ColumnIDS.D;
            //configurazione.IpotesiStudio_PrimaRigaFornitori = 4;

            //// Foglio "Lista dati"
            //configurazione.ListaDati_RigaReparti = 2;
            //configurazione.ListaDati_PrimaColonnaReparti = (int)ColumnIDS.G;
            //configurazione.ListaDati_PrimaRigaFornitori = 3;
            //configurazione.ListaDati_ColonnaSiglaFornitoriTabellaSX = (int)ColumnIDS.B;
            //configurazione.ListaDati_ColonnaCostiOrariStandard = (int)ColumnIDS.C;
            //configurazione.ListaDati_ColonnaCategoriaFornitori = (int)ColumnIDS.D;
            //configurazione.ListaDati_ColonnaSiglaFornitoriTabellaDX = (int)ColumnIDS.F;
            //configurazione.ListaDati_ColonnaReparto = (int)ColumnIDS.G;

            //// Foglio "Reportistica"
            //configurazione.Reportistica_ColonnaSigle_SX = (int)ColumnIDS.X;
            //configurazione.Reportistica_ColonnaSigle_DX = (int)ColumnIDS.AH;
            //configurazione.Reportistica_PrimaRigaFornitori = 8;
            //configurazione.Reportistica_PrimaRigaResponsabili = 11;
            //configurazione.Reportistica_ColonnaResponsabili = (int)ColumnIDS.F;
            //configurazione.Reportistica_ColonnaRestoOre = (int)ColumnIDS.P;
            //configurazione.Reportistica_ColonnaRestoEuro = (int)ColumnIDS.R;
            //configurazione.Reportistica_ColonnaRestoLump = (int)ColumnIDS.T;
            //configurazione.Reportistica_PrimaRigaNascondibile = 70;
            //configurazione.Reportistica_ValoreSegnapostoPerFormule = "999909999";
            //configurazione.Reportistica_FormulaPerColonnaRestoOre = "IF($G11<>\"non-visibile\",IFERROR(J11-999909999,0),\"non-visibile\")";
            //configurazione.Reportistica_FormulaPerColonnaRestoEuro = "IF($G11<>\"non-visibile\",IFERROR(K11-999909999,0),\"non-visibile\")";
            //configurazione.Reportistica_FormulaPerColonnaRestoLump = "IF($G11<>\"non-visibile\",IFERROR(L11-999909999,0),\"non-visibile\")";


            //// Foglio "Reportistica per tipologia"
            //configurazione.ReportisticaPerTipologia_ColonnaCategorieFornitori = (int)ColumnIDS.I;
            //configurazione.ReportisticaPerTipologia_PrimaRigaCategorieFornitori = 5;

            //configurazione.ReportisticaPerTipologia_Colonna_ConsumiSpeseAdOre_Ore = (int)ColumnIDS.J;
            //configurazione.ReportisticaPerTipologia_Formula_ConsumiSpeseAdOre_Ore = Environment.NewLine + "+IFERROR(GETPIVOTDATA(\"Actual ORE ripartito\",'Altre Tabelle di dettaglio'!$BK$5,\"Fornitore .\",\"{0}\"),0)";

            //configurazione.ReportisticaPerTipologia_Colonna_ConsumiSpeseAdOre_Euro = (int)ColumnIDS.K;
            //configurazione.ReportisticaPerTipologia_Formula_ConsumiSpeseAdOre_Euro = Environment.NewLine + "+IFERROR(GETPIVOTDATA(\"Actual EURO ripartito\",'Altre Tabelle di dettaglio'!$CH$5,\"Fornitore .\",\"{0}\"),0)";

            //configurazione.ReportisticaPerTipologia_Colonna_ConsumiSpeseLumpSum_Euro = (int)ColumnIDS.L;
            //configurazione.ReportisticaPerTipologia_Formula_ConsumiSpeseLumpSum_Euro = Environment.NewLine + "+IFERROR(GETPIVOTDATA(\"Actual LUMP ripartito\",'Altre Tabelle di dettaglio'!$DC$5,\"Fornitore .\",\"{0}\"),0)";

            //configurazione.ReportisticaPerTipologia_Colonna_AllocateSpeseAdOre_Ore = (int)ColumnIDS.N;
            //configurazione.ReportisticaPerTipologia_Formula_AllocateSpeseAdOre_Ore = Environment.NewLine + "+IFERROR(GETPIVOTDATA(\"TOTALE ORE\",'Altre Tabelle di dettaglio'!$AP$5,\"Fornitore .\",\"{0}\"),0)";

            //configurazione.ReportisticaPerTipologia_Colonna_AllocateSpeseAdOre_Euro = (int)ColumnIDS.O;
            //configurazione.ReportisticaPerTipologia_Formula_AllocateSpeseAdOre_Euro = Environment.NewLine + "+IFERROR(GETPIVOTDATA(\"BUDGET STIMATO ORE\",'Altre Tabelle di dettaglio'!$B$5,\"Fornitore .\",\"{0}\"),0)";

            //configurazione.ReportisticaPerTipologia_Colonna_AllocateSpeseLumpSum_Euro = (int)ColumnIDS.P;
            //configurazione.ReportisticaPerTipologia_Formula_AllocateSpeseLumpSum_Euro = Environment.NewLine + "+IFERROR(GETPIVOTDATA(\"BUDGET STIMATO LUMP SUM\",'Altre Tabelle di dettaglio'!$V$5,\"Fornitore .\",\"{0}\"),0)";


            //// Foglio "Sintesi"
            //configurazione.Sintesi_Riga_IntestazioneMesi = 18;
            //configurazione.Sintesi_Riga_PrimaConDati = 19;

            //configurazione.Sintesi_ColonnaSiglaFornitore = (int)ColumnIDS.A;
            //configurazione.Sintesi_ColonnaCategoriaFornitore = (int)ColumnIDS.C;
            //configurazione.Sintesi_ColonnaReparto = (int)ColumnIDS.I;
            //configurazione.Sintesi_ColonnaTipologiaSpesa = (int)ColumnIDS.K;
            //configurazione.Sintesi_Colonna_PrimaRangeMesi = (int)ColumnIDS.L; // Gen
            //configurazione.Sintesi_Colonna_UltimaRangeMesi = (int)ColumnIDS.W; // Dic
            //configurazione.Sintesi_ColonnaOverBudget= (int)ColumnIDS.X; // Presente solo in foglio gemello
            //configurazione.Sintesi_Colonna_Formula_Finire_Ore = (int)ColumnIDS.AJ;     // ORE a finire
            //configurazione.Sintesi_Colonna_Formula_Finire_LumpSum = (int)ColumnIDS.AL; // EURO LUMP a finire
            //configurazione.Sintesi_Colonna_Formula_AdOggi_Ore = (int)ColumnIDS.AM;     // ORE a oggi
            //configurazione.Sintesi_Colonna_Formula_AdOggi_LumpSum = (int)ColumnIDS.AO; // EURO LUMP a

            //configurazione.Sintesi_Formula_Finire = "IF($K19=\"{0}\",SUM({1}19:W19),0)"; // Somma valori da colonna <X> a W            
            //configurazione.Sintesi_Formula_AdOggi = "IF($K19=\"{0}\",SUM(L19:{1}19),0)"; // Somma valori da colonna L a <X-1>

            //// Foglio "Previsione a finire"
            //configurazione.PrevisioneAfinire_ColonnaRanRateOre = (int)ColumnIDS.S;
            //configurazione.PrevisioneAfinire_Riga_PrimaConDati = 6;
            //configurazione.PrevisioneAfinire_FormulaRanRateOre = "IF($G6<>\"non-visibile\",K6*12/{0},\"non-visibile\")";

            //// Fogli "Spese..."
            //configurazione.Spese_Colonna_PrimaRangeMesi = (int)ColumnIDS.D;
            //configurazione.Spese_Riga_IntestazioneMesi = 1;
            //configurazione.Spese_Riga_PrimaConDati = 2;
            //#endregion

            #region Cartella "Debug"
            configurazione.AutoSaveDebugFile = true;
            #endregion

            return configurazione;
        }
    }
}