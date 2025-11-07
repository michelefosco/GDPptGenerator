using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FilesEditor.Helpers
{
    internal class EPPlusHelperUtilities
    {
        static internal EPPlusHelper GetEPPlusHelperForExistingFile(string filePath, FileTypes fileType)
        {
            var ePPlusHelper = new EPPlusHelper();
            if (!ePPlusHelper.Open(filePath))
            {
                throw new ManagedException(
                    filePath: ePPlusHelper.FilePathInUse,
                    fileType: fileType,
                    //
                    worksheetName: null,
                    cellRow: null,
                    cellColumn: null,
                    valueHeader: ValueHeaders.None,
                    value: null,
                    //
                    errorType: ErrorTypes.UnableToOpenFile,
                    userMessage: UserErrorMessages.UnableToOpenFile
                    );
            }
            return ePPlusHelper;
        }

        static internal void ThrowExpetionsForMissingWorksheet(EPPlusHelper ePPlusHelper, string worksheetName, FileTypes fileType)
        {
            if (!ePPlusHelper.WorksheetExists(worksheetName))
            {
                throw new ManagedException(
                    filePath: ePPlusHelper.FilePathInUse,
                    fileType: fileType,
                    //
                    worksheetName: worksheetName,
                    cellRow: null,
                    cellColumn: null,
                    valueHeader: ValueHeaders.None,
                    value: null,
                    //
                    errorType: ErrorTypes.MissingWorksheet,
                    userMessage: string.Format(UserErrorMessages.MissingWorksheet, worksheetName)
                    );
            }
        }

        static internal void ThrowExpetionsForMissingHeader(EPPlusHelper ePPlusHelper, string worksheetName, FileTypes fileType, int rowWithHeaders, int headersFirstColumn, List<string> expectedColumns, string ovverideMessage = "")
        {
            var columnsList = ePPlusHelper.GetHeadersFromRow(worksheetName, rowWithHeaders, headersFirstColumn, true);
            foreach (var expectedColumn in expectedColumns)
            {
                if (!columnsList.Any(_ => _.Equals(expectedColumn, StringComparison.InvariantCultureIgnoreCase)))
                {
                    var errorMessage = string.IsNullOrEmpty(ovverideMessage)
                        ? string.Format(UserErrorMessages.MissingHeader, fileType, expectedColumn, worksheetName)
                        : ovverideMessage;

                    throw new ManagedException(
                            filePath: ePPlusHelper.FilePathInUse,
                            fileType: fileType,
                            //
                            worksheetName: worksheetName,
                            cellRow: rowWithHeaders,
                            cellColumn: null,
                            valueHeader: ValueHeaders.None,
                            value: null,
                            //
                            errorType: ErrorTypes.MissingHeader,
                            userMessage: errorMessage
                            );
                }
            }
        }
    }
}
