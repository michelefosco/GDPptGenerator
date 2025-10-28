using EPPlusExtensions;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using System.Collections.Generic;
using System.IO;
namespace FilesEditor.Helpers
{
    public class DebugInfoLogger
    {
        private class WorkSheetNames
        {
            // lunghezza massima nome foglio in Excel = 31     "1234567890123456789012345678901"
            public const string Log = "Log";
            public const string StepContext = "StepContext";
            public const string Warnings = "Warnings";
            public const string RigheInputFiles = "Righe input files";



            public const string UpdateReportsOutput = "UpdateReportsOutput";
            public const string RepartiCensitiSuController = "Reparti in controller";
            public const string RepartiCensitiSuReport = "Reparti in report";
            public const string FornitoriCensiti = "Fornitori-Censiti";
            public const string FornitoriAggiuntiDaInterfaccia = "Fornitori-Aggiunti da IU";
            public const string FornitoriAggiuntiDaListaDati = "Fornitori-Aggiun da Lista dati";
            public const string FornitoriNonCensitiInReport = "Fornitori-NON censiti in report";
            public const string Spese = "Spese";
            public const string RigheSpesaSkippate = "Righe spesa skippate";

            public const string ConsumiSpacchettati = "Consumi spacchettati";
            public const string SintesiRigheNecessarie = "Sintesi-Righe necessarie";
            public const string SintesiRigheMancanti = "Sintesi-Righe mancanti";
            public const string SintesiDatiPreElaborazione = "Sintesi-Dati PRE elaborazione";
            public const string SintesiDatiSpeseConfermate_Totali = "Sintesi-Dati spese Totali";
            public const string SintesiDatiSpeseConfermate_Actual = "Sintesi-Dati spese Actual";
            public const string SintesiDatiSpeseConfermate_Commitment = "Sintesi-Dati spese Commitment";
            public const string SintesiDatiPostElaborazione = "Sintesi-Dati POST elaborazione";
            public const string LogModificheTabellaSintesi = "Sintesi-Modifiche alla tabella";
            public const string CategorieFornitori = "Categorie fornitori";
        }

        private readonly EPPlusHelper _epPlusHelper;
        private readonly bool _autoSave;

        internal DebugInfoLogger(string filePath, bool autoSave = false)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return;

            _autoSave = autoSave;

            // Verifica che il file da generare NON esita già
            if (File.Exists(filePath))
            {
                //TEST:SituazioniNonValide.OutputFile_Debug.PercorsoFile_DebugOutPut_NON_Corretto_GiaEsistente()
                throw new ManagedException(
                    filePath: filePath,
                    fileType: FileTypes.Debug,
                    //
                    worksheetName: null,
                    cellRow: null,
                    cellColumn: null,
                    valueHeader: ValueHeaders.None,
                    value: null,
                    //
                    errorType: ErrorTypes.FileAlreadyExists,
                    userMessage: $"The file debug already exists (path: {filePath}"
                    );
            }

            _epPlusHelper = new EPPlusHelper();

            // Verifica che il file da usarsi per il debug sia stato stato creato correttamente
            if (!_epPlusHelper.Create(filePath, WorkSheetNames.Log))
            {
                throw new ManagedException(
                    filePath: filePath,
                    fileType: FileTypes.Debug,
                    //
                    worksheetName: null,
                    cellRow: null,
                    cellColumn: null,
                    valueHeader: ValueHeaders.None,
                    value: null,
                    //
                    errorType: ErrorTypes.UnableToCreateFile,
                    userMessage: $"Unable to create the file (path: {filePath}"
                    );
            }
        }

        internal void LogText(object text1, object text2 = null, object text3 = null)
        {
            if (_epPlusHelper == null) { return; }

            var worksheetName = WorkSheetNames.Log;

            _epPlusHelper.AddNewContentRow(worksheetName, text1, text2, text3);

            AutoSave();
        }

        internal void LogBuildPresentationInput(BuildPresentationInput buildPresentationInput)
        {
            if (_epPlusHelper == null) { return; }

            //   var worksheetName = WorkSheetNames.UpdateReportsInput;

            //_epPlusHelper.AddNewContentRow(worksheetName, "Periodo", updateReportsInput.Periodo);
            //_epPlusHelper.AddNewContentRow(worksheetName, "DataAggiornamento", updateReportsInput.DataAggiornamento.ToShortDateString());
            //_epPlusHelper.AddNewContentRow(worksheetName, "FileController_FilePath", updateReportsInput.FileController_FilePath);
            //_epPlusHelper.AddNewContentRow(worksheetName, "FileReport_FilePath", updateReportsInput.FileReport_FilePath);
            //_epPlusHelper.AddNewContentRow(worksheetName, "NewReport_FilePath", updateReportsInput.NewReport_FilePath);
            //_epPlusHelper.AddNewContentRow(worksheetName, "FileDebug_FilePath", updateReportsInput.FileDebug_FilePath);
            //_epPlusHelper.AddNewContentRow(worksheetName, "FornitoriDaAggiungere", "Vedo foglio dedicato");

            AutoSave();
        }

        internal void LogBuildPresentationOutput(BuildPresentationOutput buildPresentationOutput)
        {
            if (_epPlusHelper == null) { return; }

            var worksheetName = WorkSheetNames.UpdateReportsOutput;

            //if (updateReportsOutput.RepartiCensitiInController != null)
            //{
            //    _epPlusHelper.AddNewContentRow(worksheetName, "Numero RepartiCensitiInController", updateReportsOutput.RepartiCensitiInController.Count);
            //}
            //if (updateReportsOutput.RepartiCensitiInReport != null)
            //{
            //    _epPlusHelper.AddNewContentRow(worksheetName, "Numero RepartiCensitiInReport", updateReportsOutput.RepartiCensitiInReport.Count);
            //}
            //if (updateReportsOutput.FornitoriCensitiInReport != null)
            //{
            //    _epPlusHelper.AddNewContentRow(worksheetName, "Numero FornitoriCensiti", updateReportsOutput.FornitoriCensitiInReport.Count);
            //}
            //if (updateReportsOutput.RigheSpesa != null)
            //{
            //    _epPlusHelper.AddNewContentRow(worksheetName, "Numero RigheSpesa", updateReportsOutput.RigheSpesa.Count);
            //}
            //if (updateReportsOutput.RigheSpesaSkippate != null)
            //{
            //    _epPlusHelper.AddNewContentRow(worksheetName, "Numero RigheSpesaSkippate", updateReportsOutput.RigheSpesaSkippate.Count);
            //}
            //if (updateReportsOutput.RigheTabellaAvanzamento != null)
            //{
            //    _epPlusHelper.AddNewContentRow(worksheetName, "Numero RigheTabellaAvanzamento", updateReportsOutput.RigheTabellaAvanzamento.Count);
            //}
            //if (updateReportsOutput.RigheTabellaConsumiSpacchettati != null)
            //{
            //    _epPlusHelper.AddNewContentRow(worksheetName, "Numero RigheTabellaConsumiSpacchettati", updateReportsOutput.RigheTabellaConsumiSpacchettati.Count);
            //}
            //if (updateReportsOutput.RigheTabellaSintesi_SetMinimoRigheNecessarie != null)
            //{
            //    _epPlusHelper.AddNewContentRow(worksheetName, "Numero RigheTabellaSintesi", updateReportsOutput.RigheTabellaSintesi_SetMinimoRigheNecessarie.Count);
            //}
            //if (updateReportsOutput.RigheTabellaSintesi_RigheMancanti != null)
            //{
            //    _epPlusHelper.AddNewContentRow(worksheetName, "Numero RigheMancantiSuTabellaSintesi", updateReportsOutput.RigheTabellaSintesi_RigheMancanti.Count);
            //}
            //if (updateReportsOutput.CategorieFornitori != null)
            //{
            //    _epPlusHelper.AddNewContentRow(worksheetName, "Numero CategorieFornitori", updateReportsOutput.CategorieFornitori.Count);
            //}
            //if (updateReportsOutput.FornitoriNonCensitiInReport != null)
            //{
            //    _epPlusHelper.AddNewContentRow(worksheetName, "Numero FornitoriNonCensitiInReport", updateReportsOutput.FornitoriNonCensitiInReport.Count);
            //}
            if (buildPresentationOutput.ManagedException != null)
            {
                _epPlusHelper.AddNewContentRow(worksheetName, "ManagedException.InnerException.MessaggioPerUtente", buildPresentationOutput.ManagedException.UserMessage.ToString());
                _epPlusHelper.AddNewContentRow(worksheetName, "ManagedException.InnerException", buildPresentationOutput.ManagedException.ToString());
                if (buildPresentationOutput.ManagedException.InnerException != null)
                {
                    _epPlusHelper.AddNewContentRow(worksheetName, "ManagedException.InnerException", buildPresentationOutput.ManagedException.InnerException.ToString());
                }
            }

            AutoSave();
        }

        internal void LogAlias(List<AliasDefinition> aliasDefinitions, string fieldName)
        {
            if (_epPlusHelper == null) { return; }

            var worksheetName = $"Alias for {fieldName}";

            // riga intestazione
            _epPlusHelper.AddNewHeaderRow(worksheetName,
                "Raw Value",        //#1
                "Is reg. expres.",  //#2
                "New Value"         //#3
                );

            foreach (var aliasDefinition in aliasDefinitions)
            {
                _epPlusHelper.AddNewContentRow(worksheetName,
                    aliasDefinition.RawValue,               //#1
                    aliasDefinition.IsRegularExpression,    //#2
                    aliasDefinition.NewValue                //#3
                    );
            }

            AutoSave();
        }

        internal void LogRigheInputFiles(FileTypes fileType, int totRighePreservate, int totRigheEliminate, int totRigheAggiunte)
        {
            if (_epPlusHelper == null) { return; }

            var worksheetName = WorkSheetNames.RigheInputFiles;

            _epPlusHelper.AddNewContentRow(worksheetName, fileType.ToString(), "Tot righe preservate:", totRighePreservate);
            _epPlusHelper.AddNewContentRow(worksheetName, fileType.ToString(), "Tot righe eliminate:", totRigheEliminate);
            _epPlusHelper.AddNewContentRow(worksheetName, fileType.ToString(), "Tot rRighe aggiunte:", totRigheAggiunte);

            AutoSave();
        }














        //#region Reparti
        //internal void LogRepartiCensitiSuController(List<Reparto> repartiCensiti)
        //{
        //    if (_epPlusHelper == null) { return; }

        //    LogRepartiCensiti(repartiCensiti, WorkSheetNames.RepartiCensitiSuController);
        //}
        //internal void LogrepartiCensitiSuReport(List<Reparto> repartiCensiti)
        //{
        //    if (_epPlusHelper == null) { return; }

        //    LogRepartiCensiti(repartiCensiti, WorkSheetNames.RepartiCensitiSuReport);
        //}
        //private void LogRepartiCensiti(List<Reparto> repartiCensiti, string worksheetName)
        //{
        //    if (_epPlusHelper == null) { return; }

        //    // riga intestazione
        //    _epPlusHelper.AddNewHeaderRow(worksheetName,
        //        "Nome",//#1
        //        "Centro di costo",//#2
        //        "Totale centri di costo per reparto"//#3
        //        );

        //    foreach (var reparto in repartiCensiti.OrderBy(_ => _.Nome))
        //    {
        //        var primaRigaReparto = true;
        //        foreach (var centroDiCosto in reparto.CentriDiCosto.OrderBy(_ => _))
        //        {
        //            _epPlusHelper.AddNewContentRow(worksheetName,
        //                reparto.Nome,//#1
        //                primaRigaReparto ? centroDiCosto : "",//#2
        //                reparto.CentriDiCosto.Count//#3
        //                );
        //            primaRigaReparto = false;
        //        }
        //    }

        //    AutoSave();
        //}
        //#endregion



        ////internal void LogSpese(List<RigaSpese> righeSpese, string subSetName)
        ////{
        ////    if (_epPlusHelper == null) { return; }

        ////    var worksheetName = $"{WorkSheetNames.Spese}-{subSetName}";

        ////    // riga intestazione
        ////    _epPlusHelper.AddNewHeaderRow(worksheetName,
        ////        "Reparto",//#1
        ////        "Sigla fornitore",//#2
        ////        "Nome fornitore",//#3
        ////        "TipologiaDiSpesa",//#4
        ////        "StatusSpesa",//#5
        ////        "Spesa",//#6
        ////        "Ore",//#7
        ////        "CostoOrarioApplicato",//8
        ////        "Data inizio",//#9
        ////        "Data fine",//#10
        ////        "Mese inizio spalmo",//#11
        ////        "Mese fine spalmo",//#12
        ////        "Numero mesi spalmo", //13
        ////        "Spesa per mese" //14
        ////        );

        ////    // righe spese lette
        ////    foreach (var rigaSpese in righeSpese.OrderBy(_ => _.NomeReparto).ThenBy(_ => _.Fornitore.SiglaInReport))
        ////    {
        ////        _epPlusHelper.AddNewContentRow(worksheetName,
        ////            rigaSpese.NomeReparto,//#1
        ////            rigaSpese.Fornitore.SiglaInReport,//#2
        ////            rigaSpese.Fornitore.NomeSuController,//#3
        ////            rigaSpese.TipologiaDiSpesa.GetEnumDescription(),//#4
        ////            rigaSpese.StatusSpesa,//#5
        ////            rigaSpese.Spesa,//#6
        ////            rigaSpese.Ore,//#7
        ////            rigaSpese.CostoOrarioApplicato,//#8
        ////            rigaSpese.DataInizio.ToShortDateString(),//#9
        ////            rigaSpese.DataFine.ToShortDateString(),//#10
        ////            rigaSpese.DataInizioTabellaSintesi.ToString("MMM yy"), // #11
        ////            rigaSpese.DataFineTabellaSintesi.ToString("MMM yy"),// #12
        ////            rigaSpese.NumeroMesiSplitSpesaInTabellaSintesi,
        ////            (rigaSpese.Spesa != 0) ? rigaSpese.Spesa / rigaSpese.NumeroMesiSplitSpesaInTabellaSintesi : 0
        ////            );
        ////    }

        ////    AutoSave();
        ////}
        ////internal void LogRigheTabellaAvanzamento(List<RigaTabellaAvanzamento> righeTabellaAvanzamento)
        ////{
        ////    if (_epPlusHelper == null) { return; }

        ////    var worksheetName = WorkSheetNames.Avanzamento;

        ////    // riga intestazione
        ////    _epPlusHelper.AddNewHeaderRow(worksheetName,
        ////                                "Reparto",//#1
        ////                                "Fornitore.Nome",//#2
        ////                                "Fornitore.Sigla",//#3
        ////                                "Euro Act+Commit",//#4
        ////                                "ORE Act + Commit",//#5
        ////                                "EURO LUMP Act + Commit"//#6
        ////                                );
        ////    // righe tabella
        ////    foreach (var rigaTabella in righeTabellaAvanzamento)
        ////    {
        ////        _epPlusHelper.AddNewContentRow(worksheetName,
        ////                    rigaTabella.Reparto,//#1
        ////                    rigaTabella.NomeFornitore,//#2
        ////                    rigaTabella.SiglaFornitore,//#3     
        ////                    rigaTabella.TotaleSpeseAdOre_Euro,//#4
        ////                    rigaTabella.TotaleSpeseAdOre_Ore,//#5     
        ////                    rigaTabella.TotaleSpeseLumpSum//#6
        ////            );
        ////    }

        ////    AutoSave();
        ////}
        ////internal void LogRigheTabellaConsumiSpacchettati(List<RigaTabellaConsumiSpacchettati> righeTabellaConsumiSpacchettati)
        ////{
        ////    if (_epPlusHelper == null) { return; }

        ////    var worksheetName = WorkSheetNames.ConsumiSpacchettati;

        ////    // riga intestazione
        ////    _epPlusHelper.AddNewHeaderRow(worksheetName,
        ////                                "Reparto",//#1
        ////                                "Ore Actual",//#2
        ////                                "Euro Actual",//#3
        ////                                "Ore Committment",//#4
        ////                                "Euro Committment",//#5
        ////                                "Lump Actual(FBL3N)",//#6
        ////                                "Lump Committment(ME5A + ME2N)"//#7
        ////                                );
        ////    foreach (var rigaTabella in righeTabellaConsumiSpacchettati)
        ////    {
        ////        _epPlusHelper.AddNewContentRow(worksheetName,
        ////                    rigaTabella.Reparto,//#1
        ////                    rigaTabella.SpesaAdOre_Actual_Ore,//#2
        ////                    rigaTabella.SpesaAdOre_Actual_Euro,//#3     
        ////                    rigaTabella.SpesaAdOre_Commitment_Ore,//#4
        ////                    rigaTabella.SpesaAdOre_Commitment_Euro,//#5     
        ////                    rigaTabella.SpesaLumpSum_Actual,//#6
        ////                    rigaTabella.SpesaLumpSum_Commitment//#7
        ////            );
        ////    }

        ////    AutoSave();
        ////}

        ////#region Tabella Sintesi
        ////internal void LogRigheTabellaSintesi_Caclcolate(List<TuplaSpesePerAnno> righeTabellaSintesi)
        ////{
        ////    if (_epPlusHelper == null) { return; }

        ////    var worksheetName = WorkSheetNames.SintesiRigheNecessarie;
        ////    // riga intestazione
        ////    _epPlusHelper.AddNewHeaderRow(worksheetName,
        ////                                "Reparto",//#1
        ////                                "SiglaFornitore",//#2
        ////                                "TipologiaDiSpesa"//#3
        ////                                );

        ////    foreach (var rigaTabella in righeTabellaSintesi)
        ////    {
        ////        _epPlusHelper.AddNewContentRow(worksheetName,
        ////                    rigaTabella.Reparto,//#1
        ////                    rigaTabella.SiglaFornitore,//#2
        ////                    rigaTabella.TipologiaDiSpesa.GetEnumDescription()//#3
        ////            );
        ////    }

        ////    AutoSave();
        ////}

        ////internal void LogRigheTabellaSintesi_Mancanti(List<RigaTabellaSintesi> righeTabellaSintesi)
        ////{
        ////    LogRigheTabellaSintesi(righeTabellaSintesi, WorkSheetNames.SintesiRigheMancanti);
        ////}

        ////private void LogRigheTabellaSintesi(List<RigaTabellaSintesi> righeTabellaSintesi, string worksheetName)
        ////{
        ////    if (_epPlusHelper == null) { return; }

        ////    // riga intestazione
        ////    _epPlusHelper.AddNewHeaderRow(worksheetName,
        ////                                "Reparto",//#1
        ////                                "SiglaFornitore",//#2
        ////                                "TipologiaDiSpesa",//#3
        ////                                "Categoria Fornitori"//#4
        ////                                );

        ////    foreach (var rigaTabella in righeTabellaSintesi)
        ////    {
        ////        _epPlusHelper.AddNewContentRow(worksheetName,
        ////                    rigaTabella.Reparto,//#1
        ////                    rigaTabella.SiglaFornitore,//#2
        ////                    rigaTabella.TipologiaDiSpesa.GetEnumDescription(),//#3
        ////                    rigaTabella.CategoriaFornitori
        ////            );
        ////    }

        ////    AutoSave();
        ////}

        ////internal void LogRigheTabellaSintesi_CurrentValues(List<RigaTabellaSintesi> righeTabellaSintesi)
        ////{
        ////    if (_epPlusHelper == null) { return; }

        ////    var worksheetName = WorkSheetNames.SintesiDatiPreElaborazione;

        ////    // riga intestazione
        ////    _epPlusHelper.AddNewHeaderRow(worksheetName,
        ////                                "Reparto",//#1
        ////                                "SiglaFornitore",//#2
        ////                                "TipologiaDiSpesa",//#3
        ////                                "GEN",
        ////                                "FEB",
        ////                                "MAR",
        ////                                "APR",
        ////                                "MAG",
        ////                                "GIU",
        ////                                "LUG",
        ////                                "AGO",
        ////                                "SET",
        ////                                "OTT",
        ////                                "NOV",
        ////                                "DIC"
        ////                                );

        ////    foreach (var rigaTabella in righeTabellaSintesi)
        ////    {
        ////        _epPlusHelper.AddNewContentRow(worksheetName,
        ////                    rigaTabella.Reparto,//#1
        ////                    rigaTabella.SiglaFornitore,//#2
        ////                    rigaTabella.TipologiaDiSpesa.GetEnumDescription(),//#3
        ////                    rigaTabella.BudgetStanziato[0],
        ////                    rigaTabella.BudgetStanziato[1],
        ////                    rigaTabella.BudgetStanziato[2],
        ////                    rigaTabella.BudgetStanziato[3],
        ////                    rigaTabella.BudgetStanziato[4],
        ////                    rigaTabella.BudgetStanziato[5],
        ////                    rigaTabella.BudgetStanziato[6],
        ////                    rigaTabella.BudgetStanziato[7],
        ////                    rigaTabella.BudgetStanziato[8],
        ////                    rigaTabella.BudgetStanziato[9],
        ////                    rigaTabella.BudgetStanziato[10],
        ////                    rigaTabella.BudgetStanziato[11]
        ////            );
        ////    }

        ////    AutoSave();
        ////}

        ////internal void LogRigheTabellaSintesi_DatiSpeseConfermate(List<TuplaSpesePerAnno> righeTabellaSintesi)
        ////{
        ////    LogRigheTabellaSintesi_DatiSpeseConfermateTotali(righeTabellaSintesi);
        ////    LogRigheTabellaSintesi_DatiSpeseConfermateActual(righeTabellaSintesi);
        ////    LogRigheTabellaSintesi_DatiSpeseConfermateCommitment(righeTabellaSintesi);
        ////}

        //private void LogRigheTabellaSintesi_DatiSpeseConfermateTotali(List<TuplaSpesePerAnno> righeTabellaSintesi)
        //{
        //    if (_epPlusHelper == null) { return; }

        //    var worksheetName = WorkSheetNames.SintesiDatiSpeseConfermate_Totali;

        //    // riga intestazione
        //    _epPlusHelper.AddNewHeaderRow(worksheetName,
        //                                "Reparto",//#1
        //                                "SiglaFornitore",//#2
        //                                "TipologiaDiSpesa",//#3
        //                                "GEN",
        //                                "FEB",
        //                                "MAR",
        //                                "APR",
        //                                "MAG",
        //                                "GIU",
        //                                "LUG",
        //                                "AGO",
        //                                "SET",
        //                                "OTT",
        //                                "NOV",
        //                                "DIC"
        //                                );

        //    foreach (var rigaTabella in righeTabellaSintesi)
        //    {
        //        _epPlusHelper.AddNewContentRow(worksheetName,
        //                    rigaTabella.Reparto,//#1
        //                    rigaTabella.SiglaFornitore,//#2
        //                    rigaTabella.TipologiaDiSpesa.GetEnumDescription(),//#3
        //                    rigaTabella.SpeseCommittedPerMese[0] + rigaTabella.SpeseActualPerMese[0],
        //                    rigaTabella.SpeseCommittedPerMese[1] + rigaTabella.SpeseActualPerMese[1],
        //                    rigaTabella.SpeseCommittedPerMese[2] + rigaTabella.SpeseActualPerMese[2],
        //                    rigaTabella.SpeseCommittedPerMese[3] + rigaTabella.SpeseActualPerMese[3],
        //                    rigaTabella.SpeseCommittedPerMese[4] + rigaTabella.SpeseActualPerMese[4],
        //                    rigaTabella.SpeseCommittedPerMese[5] + rigaTabella.SpeseActualPerMese[5],
        //                    rigaTabella.SpeseCommittedPerMese[6] + rigaTabella.SpeseActualPerMese[6],
        //                    rigaTabella.SpeseCommittedPerMese[7] + rigaTabella.SpeseActualPerMese[7],
        //                    rigaTabella.SpeseCommittedPerMese[8] + rigaTabella.SpeseActualPerMese[8],
        //                    rigaTabella.SpeseCommittedPerMese[9] + rigaTabella.SpeseActualPerMese[9],
        //                    rigaTabella.SpeseCommittedPerMese[10] + rigaTabella.SpeseActualPerMese[10],
        //                    rigaTabella.SpeseCommittedPerMese[11] + rigaTabella.SpeseActualPerMese[11]
        //            );
        //    }

        //    AutoSave();
        //}

        //private void LogRigheTabellaSintesi_DatiSpeseConfermateActual(List<TuplaSpesePerAnno> righeTabellaSintesi)
        //{
        //    if (_epPlusHelper == null) { return; }

        //    var worksheetName = WorkSheetNames.SintesiDatiSpeseConfermate_Actual;

        //    // riga intestazione
        //    _epPlusHelper.AddNewHeaderRow(worksheetName,
        //                                "Reparto",//#1
        //                                "SiglaFornitore",//#2
        //                                "TipologiaDiSpesa",//#3
        //                                "GEN",
        //                                "FEB",
        //                                "MAR",
        //                                "APR",
        //                                "MAG",
        //                                "GIU",
        //                                "LUG",
        //                                "AGO",
        //                                "SET",
        //                                "OTT",
        //                                "NOV",
        //                                "DIC"
        //                                );

        //    foreach (var rigaTabella in righeTabellaSintesi)
        //    {
        //        _epPlusHelper.AddNewContentRow(worksheetName,
        //                    rigaTabella.Reparto,//#1
        //                    rigaTabella.SiglaFornitore,//#2
        //                    rigaTabella.TipologiaDiSpesa.GetEnumDescription(),//#3
        //                    rigaTabella.SpeseActualPerMese[0],
        //                    rigaTabella.SpeseActualPerMese[1],
        //                    rigaTabella.SpeseActualPerMese[2],
        //                    rigaTabella.SpeseActualPerMese[3],
        //                    rigaTabella.SpeseActualPerMese[4],
        //                    rigaTabella.SpeseActualPerMese[5],
        //                    rigaTabella.SpeseActualPerMese[6],
        //                    rigaTabella.SpeseActualPerMese[7],
        //                    rigaTabella.SpeseActualPerMese[8],
        //                    rigaTabella.SpeseActualPerMese[9],
        //                    rigaTabella.SpeseActualPerMese[10],
        //                    rigaTabella.SpeseActualPerMese[11]
        //            );
        //    }

        //    AutoSave();
        //}

        //private void LogRigheTabellaSintesi_DatiSpeseConfermateCommitment(List<TuplaSpesePerAnno> righeTabellaSintesi)
        //{
        //    if (_epPlusHelper == null) { return; }

        //    var worksheetName = WorkSheetNames.SintesiDatiSpeseConfermate_Commitment;

        //    // riga intestazione
        //    _epPlusHelper.AddNewHeaderRow(worksheetName,
        //                                "Reparto",//#1
        //                                "SiglaFornitore",//#2
        //                                "TipologiaDiSpesa",//#3
        //                                "GEN",
        //                                "FEB",
        //                                "MAR",
        //                                "APR",
        //                                "MAG",
        //                                "GIU",
        //                                "LUG",
        //                                "AGO",
        //                                "SET",
        //                                "OTT",
        //                                "NOV",
        //                                "DIC"
        //                                );

        //    foreach (var rigaTabella in righeTabellaSintesi)
        //    {
        //        _epPlusHelper.AddNewContentRow(worksheetName,
        //                    rigaTabella.Reparto,//#1
        //                    rigaTabella.SiglaFornitore,//#2
        //                    rigaTabella.TipologiaDiSpesa.GetEnumDescription(),//#3
        //                    rigaTabella.SpeseCommittedPerMese[0],
        //                    rigaTabella.SpeseCommittedPerMese[1],
        //                    rigaTabella.SpeseCommittedPerMese[2],
        //                    rigaTabella.SpeseCommittedPerMese[3],
        //                    rigaTabella.SpeseCommittedPerMese[4],
        //                    rigaTabella.SpeseCommittedPerMese[5],
        //                    rigaTabella.SpeseCommittedPerMese[6],
        //                    rigaTabella.SpeseCommittedPerMese[7],
        //                    rigaTabella.SpeseCommittedPerMese[8],
        //                    rigaTabella.SpeseCommittedPerMese[9],
        //                    rigaTabella.SpeseCommittedPerMese[10],
        //                    rigaTabella.SpeseCommittedPerMese[11]
        //            );
        //    }

        //    AutoSave();
        //}

        //internal void LogRigheTabellaSintesi_NewValues(List<RigaTabellaSintesi> righeTabellaSintesi)
        //{
        //    if (_epPlusHelper == null) { return; }

        //    var worksheetName = WorkSheetNames.SintesiDatiPostElaborazione;

        //    // riga intestazione
        //    _epPlusHelper.AddNewHeaderRow(worksheetName,
        //                                "Reparto",//#1
        //                                "SiglaFornitore",//#2
        //                                "TipologiaDiSpesa",//#3
        //                                "GEN",
        //                                "FEB",
        //                                "MAR",
        //                                "APR",
        //                                "MAG",
        //                                "GIU",
        //                                "LUG",
        //                                "AGO",
        //                                "SET",
        //                                "OTT",
        //                                "NOV",
        //                                "DIC"
        //                                );

        //    foreach (var rigaTabella in righeTabellaSintesi)
        //    {
        //        _epPlusHelper.AddNewContentRow(worksheetName,
        //                    rigaTabella.Reparto,//#1
        //                    rigaTabella.SiglaFornitore,//#2
        //                    rigaTabella.TipologiaDiSpesa.GetEnumDescription(),//#3
        //                    rigaTabella.ValorePianificazioneConConsumi[0],
        //                    rigaTabella.ValorePianificazioneConConsumi[1],
        //                    rigaTabella.ValorePianificazioneConConsumi[2],
        //                    rigaTabella.ValorePianificazioneConConsumi[3],
        //                    rigaTabella.ValorePianificazioneConConsumi[4],
        //                    rigaTabella.ValorePianificazioneConConsumi[5],
        //                    rigaTabella.ValorePianificazioneConConsumi[6],
        //                    rigaTabella.ValorePianificazioneConConsumi[7],
        //                    rigaTabella.ValorePianificazioneConConsumi[8],
        //                    rigaTabella.ValorePianificazioneConConsumi[9],
        //                    rigaTabella.ValorePianificazioneConConsumi[10],
        //                    rigaTabella.ValorePianificazioneConConsumi[11]
        //            );
        //    }

        //    AutoSave();
        //}
        //#endregion

        internal void LogCategorieFornitori(List<string> categorieFornitori)
        {
            if (_epPlusHelper == null) { return; }

            var worksheetName = WorkSheetNames.CategorieFornitori;

            // riga intestazione
            _epPlusHelper.AddNewHeaderRow(worksheetName,
                "Categoria fornitore"//#1
                );

            foreach (var categoriaFornitori in categorieFornitori)
            {
                _epPlusHelper.AddNewContentRow(worksheetName,
                   categoriaFornitori//#1 
                    );
            }

            AutoSave();
        }




        //internal void LogFormuleReportisticaPerCategoriaIntestazione()
        //{
        //    if (_epPlusHelper == null) { return; }

        //    var worksheetName = WorkSheetNames.FormuleReportisticaPerCategoria;

        //    // riga intestazione
        //    _epPlusHelper.AddNewHeaderRow(worksheetName,
        //        "Categoria",//#1
        //        "Colonna",//#2
        //        "Porzione di formula",//#3
        //        "Formula con separatore ','",//#4
        //        "Formula con separatore ';'"//#5
        //        );

        //    AutoSave();
        //}

        //internal void LogFormuleReportisticaPerCategoriaRiga(string categoria, string colonna, string formula, int porzioneDiFormula)
        //{
        //    if (_epPlusHelper == null) { return; }

        //    var worksheetName = WorkSheetNames.FormuleReportisticaPerCategoria;

        //    _epPlusHelper.AddNewContentRow(worksheetName,
        //        categoria,//#1
        //        colonna,//#2
        //        porzioneDiFormula,//#3
        //        formula,//#4
        //        formula.Replace(",", ";")//#5
        //        );

        //    AutoSave();
        //}

        //internal void LogOperazioniSuPianificazioneConConsumi(List<RigaLogElaborazioniTabellaSintesi> operazioniSuPianificazioneConConsumi)
        //{
        //    if (_epPlusHelper == null) { return; }

        //    var worksheetName = WorkSheetNames.LogModificheTabellaSintesi;

        //    // riga intestazione
        //    _epPlusHelper.AddNewHeaderRow(worksheetName,
        //        "Reparto",//#1
        //        "Operazione",//#2
        //        "Quantità",//#3
        //        "Mese",//#4
        //        "Fornitore",//#5
        //        "Riga" //#6
        //        );

        //    foreach (var operazione in operazioniSuPianificazioneConConsumi.OrderBy(_ => _.Reparto)
        //            .ThenBy(_ => _.Mese)
        //            .ThenBy(_ => _.Fornitore)
        //            .ThenBy(_ => _.SimboloValore)
        //            .ThenBy(_ => _.Operazione))
        //    {
        //        _epPlusHelper.AddNewContentRow(worksheetName,
        //           operazione.Reparto,       //#1 
        //           operazione.Operazione,    //#2
        //           Math.Round(operazione.Valore, Numbers.NumeroDecimaliImportiSpese) + " " + operazione.SimboloValore,    //#3
        //           new DateTime(2024, operazione.Mese, 1).ToString("MMMM"),          //#4
        //           operazione.Fornitore,     //#5
        //            operazione.Riga          //#6
        //            );
        //    }


        //    // registrazione log in formato testuale
        //    //worksheetName = WorkSheetNames.LogModificheTabellaSintesi + "_V2";
        //    //var lastReparto = "";
        //    //foreach (var cambiamento in cambiamentiAllaTabellaSintesi.OrderBy(_ => _.Reparto)
        //    //        .ThenBy(_ => _.Mese)
        //    //        .ThenBy(_ => _.Fornitore)
        //    //        .ThenBy(_ => _.SimboloValore)
        //    //        .ThenBy(_ => _.Operazione))
        //    //{
        //    //    if (!cambiamento.Reparto.Equals(lastReparto))
        //    //    {
        //    //        //Aggiungo una riga di log ad ogni variazione di reparto
        //    //        lastReparto = cambiamento.Reparto;
        //    //        var testoReparto = $"Reparto {cambiamento.Reparto}";
        //    //        _epPlusHelper.AddNewContentRow(worksheetName, lastReparto);
        //    //    }

        //    //    var meseTestuale = new DateTime(2024, cambiamento.Mese, 1).ToString("MMMM");
        //    //    var testoOperazione = $"{cambiamento.Operazione} {cambiamento.Valore} {cambiamento.SimboloValore} di pianificato nel mese di {meseTestuale} per il fornitore {cambiamento.Fornitore} (riga {cambiamento.Riga})";
        //    //    _epPlusHelper.AddNewContentRow(worksheetName, "", testoOperazione);
        //    //}

        //    AutoSave();
        //}


        internal void Beautify()
        {
            if (_epPlusHelper == null) { return; }

            var worksheetNames = _epPlusHelper.GetWorksheetNames();
            foreach (var worksheetName in worksheetNames)
            {
                _epPlusHelper.BorderAllContent(worksheetName);
                _epPlusHelper.AutoFitColumns(worksheetName);
            }

            AutoSave();
        }

        internal void Save()
        {
            if (_epPlusHelper == null) { return; }

            _epPlusHelper.Save();
        }

        private void AutoSave()
        {
            if (_autoSave)
            {
                _epPlusHelper.Save();
            }
        }

        internal void LogWarning(string warningMessage)
        {
            if (_epPlusHelper == null) { return; }

            var worksheetName = WorkSheetNames.Warnings;

            _epPlusHelper.AddNewContentRow(worksheetName, warningMessage);
        }


        internal void LogStepContext(string stepName, StepContext context)
        {
            if (_epPlusHelper == null) { return; }

            var worksheetName = WorkSheetNames.StepContext;

            //todo: aggiungere tutte le properties di StepContext
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "Esito", context.Esito);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "DestinationFolder", context.DestinationFolder);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "TmpFolder", context.TmpFolder);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "DataSourceFilePath", context.DataSourceFilePath);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "DebugFilePath", context.DebugFilePath);
            //
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "ReplaceAllData_FileSuperDettagli", context.ReplaceAllData_FileSuperDettagli);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "PeriodDate", context.PeriodDate);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "FileBudgetPath", context.FileBudgetPath);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "FileForecastPath", context.FileForecastPath);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "FileSuperDettagliPath", context.FileSuperDettagliPath);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "FileRunRatePath", context.FileRunRatePath);
            //
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "Warnings (count)", context.Warnings.Count);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "Applicablefilters (count)", context.ApplicableFilters.Count);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "AliasDefinitions_BusinessTMP (count)", context.AliasDefinitions_BusinessTMP.Count);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "AliasDefinitions_Categoria (count)", context.AliasDefinitions_Categoria.Count);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "SildeToGenerate (count)", context.SildeToGenerate.Count);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "ItemsToExportAsImage (count)", context.ItemsToExportAsImage.Count);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "OutputFilePathLists (count)", context.OutputFilePathLists.Count);

            AutoSave();
        }
    }
}