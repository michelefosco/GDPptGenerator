using EPPlusExtensions;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using System;
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
            public const string RigheSourceFiles = "Righe input files";
            public const string ImageExtraction = "ImageExtraction";


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

        private string TimeStampString { get { return DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.fff"); } }


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

            _epPlusHelper.AddNewContentRow(worksheetName, TimeStampString, text1, text2, text3);

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

        internal void LogRigheSourceFiles(FileTypes fileType, int totRighePreservate, int totRigheEliminate, int totRigheAggiunte)
        {
            if (_epPlusHelper == null) { return; }

            var worksheetName = WorkSheetNames.RigheSourceFiles;

            _epPlusHelper.AddNewContentRow(worksheetName, fileType.ToString(), "Tot righe preservate:", totRighePreservate);
            _epPlusHelper.AddNewContentRow(worksheetName, fileType.ToString(), "Tot righe eliminate:", totRigheEliminate);
            _epPlusHelper.AddNewContentRow(worksheetName, fileType.ToString(), "Tot rRighe aggiunte:", totRigheAggiunte);

            AutoSave();
        }

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
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "Time stamp", TimeStampString);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "Esito", context.Esito);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "DestinationFolder", context.DestinationFolder);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "TmpFolder", context.TmpFolder);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "DataSourceFilePath", context.DataSourceFilePath);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "DebugFilePath", context.DebugFilePath);
            //
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "AppendCurrentYear_FileSuperDettagli", context.AppendCurrentYear_FileSuperDettagli);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "PeriodDate", context.PeriodDate);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "FileBudgetPath", context.FileBudgetPath);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "FileForecastPath", context.FileForecastPath);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "FileSuperDettagliPath", context.FileSuperDettagliPath);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "FileRunRatePath", context.FileRunRatePath);
            //
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "Warnings (count)", context.Warnings.Count);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "Applicablefilters (count)", context.ApplicableFilters.Count);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "AliasDefinitions_Business (count)", context.AliasDefinitions_Business.Count);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "AliasDefinitions_Categoria (count)", context.AliasDefinitions_Categoria.Count);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "SildeToGenerate (count)", context.SildeToGenerate.Count);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "ItemsToExportAsImage (count)", context.ItemsToExportAsImage.Count);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "OutputFilePathLists (count)", context.OutputFilePathLists.Count);

            AutoSave();
        }

        internal void LogInfoExportImages(int attemptNumber, string imageId, string imageFilePath, string workSheetName, string printArea, bool isPresentOnFileSistem)
        {
            if (_epPlusHelper == null) { return; }

            var worksheetName = WorkSheetNames.ImageExtraction;

            _epPlusHelper.AddNewContentRow(worksheetName, 
                TimeStampString, 
                "Attempt number:", attemptNumber, 
                "Imange info:", imageId, imageFilePath, workSheetName, printArea, 
                "Success?:", isPresentOnFileSistem);
        }
    }
}