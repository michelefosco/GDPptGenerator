using EPPlusExtensions;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using System;
using System.Collections.Generic;
using System.IO;

namespace FilesEditor.Helpers
{
    public class DebugInfoLoggerBase
    {
        const string WORKSHEET_NAME_TEXT = "Text";
        const string WORKSHEET_NAME_WARNINGS = "Warnings";
        const string WORKSHEET_NAME_PERFORMANCE = "Performance";


        internal readonly EPPlusHelper _epPlusHelper;
        private readonly bool _autoSave;

        internal string TimeStampString { get { return DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.fff"); } }


        public DebugInfoLoggerBase(string filePath, bool autoSave = false)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            { return; }

            _autoSave = autoSave;

            // Verifica che il file da generare NON esita già
            if (File.Exists(filePath))
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
                    errorType: ErrorTypes.FileAlreadyExists,
                    userMessage: $"The file debug already exists (path: {filePath}"
                    );
            }

            _epPlusHelper = new EPPlusHelper();

            // Verifica che il file da usarsi per il debug sia stato stato creato correttamente
            if (!_epPlusHelper.Create(filePath, WORKSHEET_NAME_TEXT))
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


        internal void RunAutoSave()
        {
            if (_autoSave)
            {
                _epPlusHelper.Save();
            }
        }
        public void Save()
        {
            if (_epPlusHelper == null) { return; }

            _epPlusHelper.Save();
        }

        public void LogText(object text1, object text2 = null, object text3 = null)
        {
            if (_epPlusHelper == null) { return; }

            var worksheetName = WORKSHEET_NAME_TEXT;

            _epPlusHelper.AddNewContentRow(worksheetName, TimeStampString, text1, text2, text3);

            RunAutoSave();
        }
        public void LogWarning(string warningText)
        {
            if (_epPlusHelper == null) { return; }

            var worksheetName = WORKSHEET_NAME_WARNINGS;

            _epPlusHelper.AddNewContentRow(worksheetName, TimeStampString, warningText);

            RunAutoSave();
        }

        public void LogPerformance(string taskReference, TimeSpan spentTime)
        {
            if (_epPlusHelper == null) { return; }

            var worksheetName = WORKSHEET_NAME_PERFORMANCE;

            _epPlusHelper.AddNewContentRow(worksheetName, taskReference, 
                "Totale seconds:", spentTime.TotalSeconds,
                "Total milliseconds:", spentTime.TotalMilliseconds);

            RunAutoSave();
        }

        public void Beautify()
        {
            if (_epPlusHelper == null) { return; }

            var worksheetNames = _epPlusHelper.GetWorksheetNames();
            foreach (var worksheetName in worksheetNames)
            {
                _epPlusHelper.BorderAllContent(worksheetName);
                _epPlusHelper.AutoFitColumns(worksheetName);
            }

            RunAutoSave();
        }
    }


    public class DebugInfoLogger : DebugInfoLoggerBase
    {
        private class WorkSheetNames
        {
            // lunghezza massima nome foglio in Excel = 31     "1234567890123456789012345678901"

            public const string StepContext = "StepContext";
            public const string RigheSourceFiles = "Righe input files";
            public const string ImageExtraction = "ImageExtraction";
        }

        public DebugInfoLogger(string filePath, bool autoSave = false) : base(filePath, autoSave)
        { }

        // Del dominio specifico
        public void LogAlias(List<AliasDefinition> aliasDefinitions, string fieldName)
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

            RunAutoSave();
        }
        public void LogRigheSourceFiles(FileTypes fileType, int totRighePreservate, int totRigheEliminate, int totRigheAggiunte)
        {
            if (_epPlusHelper == null) { return; }

            var worksheetName = WorkSheetNames.RigheSourceFiles;

            _epPlusHelper.AddNewContentRow(worksheetName, fileType.ToString(), "Tot righe preservate:", totRighePreservate);
            _epPlusHelper.AddNewContentRow(worksheetName, fileType.ToString(), "Tot righe eliminate:", totRigheEliminate);
            _epPlusHelper.AddNewContentRow(worksheetName, fileType.ToString(), "Tot rRighe aggiunte:", totRigheAggiunte);

            RunAutoSave();
        }
        public void LogStepContext(string stepName, StepContext context)
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

            RunAutoSave();
        }
        public void LogInfoExportImages(int attemptNumber, string imageId, string imageFilePath, string workSheetName, string printArea, bool isPresentOnFileSistem, TimeSpan timeSpent)
        {
            if (_epPlusHelper == null) { return; }

            var worksheetName = WorkSheetNames.ImageExtraction;

            _epPlusHelper.AddNewContentRow(worksheetName,
                            TimeStampString,
                            "Image info:", imageId, imageFilePath, workSheetName, printArea,
                            "Attempt number:", attemptNumber,
                            "Success?:", isPresentOnFileSistem,
                            "Milliseconds spent:", timeSpent.TotalMilliseconds
                            );
        }
    }
}