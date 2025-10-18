using DocumentFormat.OpenXml.Presentation;
using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Enums;
using System;

namespace FilesEditor.Entities.Exceptions
{
    public class ManagedException : Exception
    {
        public readonly string FilePath;
        public readonly FileTypes FileType;
        //
        public readonly string WorksheetName;
        public readonly int? CellRow;
        public readonly int? CellColumn;
        public readonly ValueHeaders ValueHeader;
        public readonly string Value;
        //
        public readonly ErrorTypes ErrorType;
        public readonly string UserMessage;

        public ManagedException(
                string filePath,
                FileTypes fileType,
                //
                string worksheetName,
                int? cellRow,
                int? cellColumn,
                ValueHeaders valueHeader,
                string value,
                //
                ErrorTypes errorType,
                string userMessage
                ) : base(userMessage)
        {

            FilePath = filePath;
            FileType = fileType;
            //
            WorksheetName = worksheetName;
            CellRow = cellRow;
            CellColumn = cellColumn;
            this.ValueHeader = valueHeader;
            Value = value;
            //
            ErrorType = errorType;
            UserMessage = string.IsNullOrEmpty(userMessage)
                        ? buildUserMessage()
                        : userMessage;
        }
        private string buildUserMessage()
        {
            var _userMessage = $"{ErrorType.GetEnumDescription()}, file type: {FileType.GetEnumDescription()}";

            if (!string.IsNullOrEmpty(WorksheetName))
            {
                _userMessage += $", worksheet: \"{WorksheetName}\"";
            }

            if (CellColumn.HasValue && CellRow.HasValue)
            {
                _userMessage += $", cell: {(ColumnIDS)CellColumn.Value}{CellRow}";
            }
            else
            {
                if (CellColumn.HasValue)
                {
                    _userMessage += $", column: {(ColumnIDS)CellColumn.Value}";
                }

                if (CellRow.HasValue)
                {
                    _userMessage += $", row: {CellRow}";
                }
            }

            if (!string.IsNullOrEmpty(Value))
            {
                _userMessage += $", value: \"{Value}\"";
            }

            if (ValueHeader != ValueHeaders.None)
            {
                _userMessage += $", value types: {ValueHeader.GetEnumDescription()}";
            }

            return _userMessage;
        }


        public ManagedException(Exception ex)
        {
            var managedException = new ManagedException(
                    filePath: null,
                    fileType: FileTypes.Undefined,
                    //
                    worksheetName: null,
                    cellRow: null,
                    cellColumn: null,
                    valueHeader: ValueHeaders.None,
                    value: null,
                    //
                    errorType: ErrorTypes.UnhandledException,
                    userMessage: ex.Message + (ex.InnerException != null ? " (" + ex.InnerException.Message + ")" : "")
                    );
        }


        internal static void ThrowIfMissingMandatoryValue(string valueToCheck, string filePath, FileTypes fileType, string worksheetName, int cellRow, int cellColumn, ValueHeaders valueHeader)
        {
            if (valueToCheck == null)
            {
                throw new ManagedException(
                    filePath: filePath,
                    fileType: fileType,
                    //
                    worksheetName: null,
                    cellRow: cellRow,
                    cellColumn: cellColumn,
                    valueHeader: valueHeader,
                    value: null,
                    //
                    errorType: ErrorTypes.MissingValue,
                    userMessage: string.Format(UserErrorMessages.MissingValue, valueHeader)
                    );
            }
        }
    }
}