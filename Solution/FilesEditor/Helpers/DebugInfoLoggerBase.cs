using EPPlusExtensions;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using System;
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
                "Seconds:", spentTime.TotalSeconds.ToString("F0"),
                "Milliseconds:", spentTime.TotalMilliseconds.ToString("F0")
                );
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
}