using FilesEditor.Entities;
using FilesEditor.Enums;
using System;
using System.Collections.Generic;

namespace FilesEditor.Helpers
{
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

        public void LogRigheSourceFiles(FileTypes fileType, InfoRows infoRows)
        {
            if (_epPlusHelper == null) { return; }

            var worksheetName = WorkSheetNames.RigheSourceFiles;

            _epPlusHelper.AddNewContentRow(worksheetName, fileType.ToString(), "Info numero righe");
            _epPlusHelper.AddNewContentRow(worksheetName, "", "Iniziali:", infoRows.Iniziali);
            _epPlusHelper.AddNewContentRow(worksheetName, "", "Preservate:", infoRows.Preservate);
            _epPlusHelper.AddNewContentRow(worksheetName, "", "Riutilizzate:", infoRows.Riutilizzate);
            _epPlusHelper.AddNewContentRow(worksheetName, "", "Eliminate:", infoRows.Eliminate);
            _epPlusHelper.AddNewContentRow(worksheetName, "", "Aggiunte:", infoRows.Aggiunte);
            _epPlusHelper.AddNewContentRow(worksheetName, "", "Finali:", infoRows.Finali);
            _epPlusHelper.AddNewContentRow(worksheetName, "---------------");
            RunAutoSave();
        }

        public void LogStepContext(string stepName, StepContext context)
        {
            if (_epPlusHelper == null) { return; }

            var worksheetName = WorkSheetNames.StepContext;

            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "Time stamp", TimeStampString);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "ElapsedTime", context.ElapsedTime.ToString());
            //
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
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "SildeToGenerate (count)", context.SlidesToGenerate.Count);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "ItemsToExportAsImage (count)", context.ItemsToExportAsImage.Count);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "OutputFilePathLists (count)", context.OutputFilePathLists.Count);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "DatasourceStatus_ImportDatiCompletato", context.DatasourceStatus_ImportDatiCompletato);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "DatasourceStatus_RefreshAllCompletato", context.DatasourceStatus_RefreshAllCompletato);
            _epPlusHelper.AddNewContentRow(worksheetName, stepName, "Esito", context.Esito);

            const string STRINGA_SEPARATORE = "--------------------";
            _epPlusHelper.AddNewContentRow(worksheetName, STRINGA_SEPARATORE, STRINGA_SEPARATORE, STRINGA_SEPARATORE);

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