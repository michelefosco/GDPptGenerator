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
            ValueHeader = valueHeader;
            Value = value;
            //
            ErrorType = errorType;
            UserMessage = string.IsNullOrEmpty(userMessage)
                        ? BuildUserMessage()
                        : userMessage;
        }
        private string BuildUserMessage()
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
            var userMessage = ex.Message;

            if (ex.InnerException != null)
                userMessage += "\r\n" +ex.InnerException.Message;

            if (ex.StackTrace != null)
                userMessage += "\r\n\r\n" + ex.StackTrace;

            FilePath = null;
            FileType = FileTypes.Undefined;
            //
            WorksheetName = null;
            CellRow = null;
            CellColumn = null;
            ValueHeader = ValueHeaders.None;
            Value = null;
            //
            ErrorType = ErrorTypes.UnhandledException;
            UserMessage = userMessage;
        }


        internal static void ThrowIfMissingMandatoryValue(string valueToCheck, string filePath, FileTypes fileType, string worksheetName, int cellRow, int cellColumn, ValueHeaders valueHeader)
        {
            if (valueToCheck == null)
            {
                throw new ManagedException(
                    filePath: filePath,
                    fileType: fileType,
                    //
                    worksheetName: worksheetName,
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