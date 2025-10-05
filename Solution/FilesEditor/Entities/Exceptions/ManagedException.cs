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
        public readonly ValueTypes ValueType;
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
                ValueTypes valueType,
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
            this.ValueType = valueType;
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

            if (ValueType != ValueTypes.None)
            {
                _userMessage += $", value types: {ValueType.GetEnumDescription()}";
            }

            return _userMessage;
        }

    }
}