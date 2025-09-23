using System.Collections.Generic;

namespace PptGenerator.Entities
{
    public class Configurazione
    {
        //#region Controller
        //// ACT SOLO CDC
        //public int ActualSoloCdc_ColonnaCentriDiCosto { get; set; }
        //public int ActualSoloCdc_ColonnaSpesa { get; set; }
        //public int ActualSoloCdc_ColonnaFornitore { get; set; }
        //public int ActualSoloCdc_ColonnaDataInizio { get; set; }
        //public int ActualSoloCdc_ColonnaDataFine { get; set; }
        //public string ActualSoloCdc_IntestazionePrimaColonna { get; set; }


        //// commitment solo cdc
        //public int CommitmentSoloCdc_ColonnaCentriDiCosto { get; set; }
        //public int CommitmentSoloCdc_ColonnaSpesa { get; set; }
        //public int CommitmentSoloCdc_ColonnaFornitore { get; set; }
        //public int CommitmentSoloCdc_ColonnaDataInizio { get; set; }
        //public int CommitmentSoloCdc_ColonnaDataFine { get; set; }
        //public string CommitmentSoloCdc_IntestazionePrimaColonna { get; set; }


        //// FBL3N_act
        //public int FBL3Nact_ColonnaCentriDiCosto { get; set; }
        //public int FBL3Nact_ColonnaSpesa { get; set; }
        //public int FBL3Nact_ColonnaFornitore { get; set; }
        //public int FBL3Nact_ColonnaDataInizio { get; set; }
        //public int FBL3Nact_ColonnaDataFine { get; set; }
        //public string FBL3Nact_IntestazionePrimaColonna { get; set; }


        //// ME5A
        //public int ME5A_ColonnaCentriDiCosto { get; set; }
        //public int ME5A_ColonnaSpesa { get; set; }
        //public int ME5A_ColonnaFornitore { get; set; }
        //public int ME5A_ColonnaDataInizio { get; set; }
        //public int ME5A_ColonnaDataFine { get; set; }
        //public string ME5A_IntestazionePrimaColonna { get; set; }

        //// ME2N
        //public int ME2N_ColonnaCentriDiCosto { get; set; }
        //public int ME2N_ColonnaSpesa { get; set; }
        //public int ME2N_ColonnaFornitore { get; set; }
        //public int ME2N_ColonnaDataInizio { get; set; }
        //public int ME2N_ColonnaDataFine { get; set; }
        //public string ME2N_IntestazionePrimaColonna { get; set; }

        //// RECAP
        //public List<string> Recap_ListaMergeReparti { get; set; }

        //#endregion


        //#region Report

        //public bool AttivazioneOpzione_RefreshOnLoad_SuTutteLePivotTable { get; set; }

        //// Anagrafica Fornitori
        //public int AnagraficaFornitori_PrimaRigaFornitori { get; set; }
        //public int AnagraficaFornitori_ColonnaSigle { get; set; }
        //public int AnagraficaFornitori_ColonnaNomi { get; set; }


        //// Avanzamento
        //public int Avanzamento_PrimaRigaDaAggiornare { get; set; }
        //public int Avanzamento_PrimaColonnaDaAggiornare { get; set; }
        //public int Avanzamento_UltimaColonnaDaAggiornare { get; set; }
        //public string Avanzamento_AddressCellaUltimoAggiornamento { get; set; }


        ////Consumi Spacchettati
        //public int ConsumiSpacchettati_PrimaRigaConDati { get; set; }
        //public int ConsumiSpacchettati_ColonnaReparti { get; set; }
        //public int ConsumiSpacchettati_ColonnaOreActual { get; set; }


        //// Ipotesi studio
        //public int IpotesiStudio_PrimaRigaFornitori { get; set; }
        //public int IpotesiStudio_ColonnaSigle { get; set; }
        //public int IpotesiStudio_ColonnaUltimaDaOrdinare { get; set; }


        //// Lista dati
        //public int ListaDati_RigaReparti { get; set; }
        //public int ListaDati_PrimaRigaFornitori { get; set; }
        //public int ListaDati_PrimaColonnaReparti { get; set; }
        //public int ListaDati_ColonnaSiglaFornitoriTabellaSX { get; set; }
        //public int ListaDati_ColonnaSiglaFornitoriTabellaDX { get; set; }
        //public int ListaDati_ColonnaCostiOrariStandard { get; set; }
        //public int ListaDati_ColonnaCategoriaFornitori { get; set; }
        //public int ListaDati_ColonnaReparto { get; set; }


        //// Recap
        //public int Recap_Riga_PrimaConDati { get; set; }
        //public int Recap_ColonnaReparti { get; set; }
        //public int Recap_ColonnaCentriDiCosto { get; set; }


        //// Reportistica
        //public int Reportistica_PrimaRigaFornitori { get; set; }
        //public int Reportistica_ColonnaSigle_SX { get; set; }
        //public int Reportistica_ColonnaSigle_DX { get; set; }
        //public int Reportistica_PrimaRigaResponsabili { get; set; }
        //public int Reportistica_ColonnaResponsabili { get; set; }
        //public int Reportistica_ColonnaRestoOre { get; set; }
        //public int Reportistica_ColonnaRestoEuro { get; set; }
        //public int Reportistica_ColonnaRestoLump { get; set; }
        //public int Reportistica_PrimaRigaNascondibile { get; set; }
        //public string Reportistica_FormulaPerColonnaRestoOre { get; set; }
        //public string Reportistica_FormulaPerColonnaRestoEuro { get; set; }
        //public string Reportistica_FormulaPerColonnaRestoLump { get; set; }
        //public string Reportistica_ValoreSegnapostoPerFormule { get; set; }


        //// Reportistica Per Tipologia
        //public int ReportisticaPerTipologia_ColonnaCategorieFornitori { get; set; }
        //public int ReportisticaPerTipologia_PrimaRigaCategorieFornitori { get; set; }

        ///// <summary>
        ///// J
        ///// </summary>
        //public int ReportisticaPerTipologia_Colonna_ConsumiSpeseAdOre_Ore { get; set; }
        ///// <summary>
        ///// J
        ///// </summary>
        //public string ReportisticaPerTipologia_Formula_ConsumiSpeseAdOre_Ore { get; set; }

        ///// <summary>
        ///// K
        ///// </summary>
        //public int ReportisticaPerTipologia_Colonna_ConsumiSpeseAdOre_Euro { get; set; }
        ///// <summary>
        ///// K
        ///// </summary>
        //public string ReportisticaPerTipologia_Formula_ConsumiSpeseAdOre_Euro { get; set; }

        ///// <summary>
        ///// L
        ///// </summary>
        //public int ReportisticaPerTipologia_Colonna_ConsumiSpeseLumpSum_Euro { get; set; }
        ///// <summary>
        ///// L
        ///// </summary>
        //public string ReportisticaPerTipologia_Formula_ConsumiSpeseLumpSum_Euro { get; set; }

        ///// <summary>
        ///// N
        ///// </summary>
        //public int ReportisticaPerTipologia_Colonna_AllocateSpeseAdOre_Ore { get; set; }
        ///// <summary>
        ///// N
        ///// </summary>
        //public string ReportisticaPerTipologia_Formula_AllocateSpeseAdOre_Ore { get; set; }

        ///// <summary>
        ///// O
        ///// </summary>
        //public int ReportisticaPerTipologia_Colonna_AllocateSpeseAdOre_Euro { get; set; }
        ///// <summary>
        ///// O
        ///// </summary>
        //public string ReportisticaPerTipologia_Formula_AllocateSpeseAdOre_Euro { get; set; }

        ///// <summary>
        ///// P
        ///// </summary>
        //public int ReportisticaPerTipologia_Colonna_AllocateSpeseLumpSum_Euro { get; set; }
        ///// <summary>
        ///// P
        ///// </summary>
        //public string ReportisticaPerTipologia_Formula_AllocateSpeseLumpSum_Euro { get; set; }


        //// Sintesi
        //public int Sintesi_Riga_PrimaConDati { get; set; }
        //public int Sintesi_Riga_IntestazioneMesi { get; set; }

        //public int Sintesi_ColonnaSiglaFornitore { get; set; }
        //public int Sintesi_ColonnaCategoriaFornitore { get; set; }
        //public int Sintesi_ColonnaReparto { get; set; }
        //public int Sintesi_ColonnaTipologiaSpesa { get; set; }
        //public int Sintesi_Colonna_PrimaRangeMesi { get; set; }
        //public int Sintesi_Colonna_UltimaRangeMesi { get; set; }
        //public int Sintesi_ColonnaOverBudget { get; set; } // Presente solo il foglio gemello
        //public int Sintesi_Colonna_Formula_Finire_Ore { get; set; }
        //public int Sintesi_Colonna_Formula_Finire_LumpSum { get; set; }
        //public int Sintesi_Colonna_Formula_AdOggi_Ore { get; set; }
        //public int Sintesi_Colonna_Formula_AdOggi_LumpSum { get; set; }

        //public string Sintesi_Formula_Finire { get; set; }
        //public string Sintesi_Formula_AdOggi { get; set; }


        //// Previsione a finire
        //public int PrevisioneAfinire_ColonnaRanRateOre { get; set; }
        //public int PrevisioneAfinire_Riga_PrimaConDati { get; set; }
        //public string PrevisioneAfinire_FormulaRanRateOre { get; set; }


        //// Fogli Spesa...
        //public int Spese_Colonna_PrimaRangeMesi { get; set; }

        //public int Spese_Riga_IntestazioneMesi { get; set; }
        //public int Spese_Riga_PrimaConDati { get; set; }
        //#endregion


        #region DebugFile
        // Salva il file debug dopo ogni scrittura
        public bool AutoSaveDebugFile { get; set; }

        #endregion
    }
}