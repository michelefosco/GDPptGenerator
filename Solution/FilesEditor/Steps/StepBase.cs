using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FilesEditor.Steps
{
    //todo: rendere internal ma accessibile ai test
    abstract public class StepBase
    {
        internal StepContext Context;

        internal StepBase(StepContext context)
        {
            Context = context;
        }

        internal abstract EsitiFinali DoSpecificTask();

        internal EsitiFinali Do()
        {
            return DoSpecificTask();
        }


        #region Utilities
        internal void CancellaFileSeEsiste(string filePath, FileTypes fileType)
        {
            if (File.Exists(filePath))
            {
                try
                {
                    File.Delete(filePath);
                }
                catch (Exception ex)
                {
                    throw new ManagedException(
                        filePath: filePath,
                        fileType: fileType,
                        //
                        worksheetName: null,
                        cellRow: null,
                        cellColumn: null,
                        valueHeader: ValueHeaders.None,
                        value: null,
                        //
                        errorType: ErrorTypes.UnableToDeleteFile,
                        userMessage: string.Format(UserErrorMessages.UnableToDeleteFile, filePath)
                        );
                }
            }
        }

        internal void CancellaDirectorySeEsiste(string folderPath)
        {
            if (Directory.Exists(folderPath))
            {
                try
                {
                    Directory.Delete(folderPath, true);
                }
                catch (Exception ex)
                {
                    throw new ManagedException(
                        filePath: folderPath,
                        fileType: FileTypes.Directory,
                        //
                        worksheetName: null,
                        cellRow: null,
                        cellColumn: null,
                        valueHeader: ValueHeaders.None,
                        value: null,
                        //
                        errorType: ErrorTypes.UnableToDeleteFile,
                        userMessage: string.Format(UserErrorMessages.UnableToDeleteFile, folderPath)
                        );
                }
            }
        }

        internal void AddWarning(string warningMessage)
        {
            Context.Warnings.Add(warningMessage);
            Context.DebugInfoLogger?.LogWarning(warningMessage);
        }

        /// <summary>
        /// Get all files matching the criteria
        /// </summary>
        /// <param name="folderPath"></param>
        /// <param name="filter"></param>
        /// <returns></returns>
        internal List<string> GetFilesListFromFolder(string folderPath, string filter)
        {
            // rimuovo dalla lista i file il cui nome inizia con "~$" (ovvero i file temporaranei creati da Excel quando un file è aperto)
            var filePaths = Directory.GetFiles(folderPath, filter, SearchOption.TopDirectoryOnly)
                    .Where(_ => !Path.GetFileName(_).StartsWith("~$")).ToList();
            return filePaths;
        }
        internal string GetImagePath(string imageId)
        {
            //todo: rivedere questo metodo. non può dipendere dal context
            var imagePath = $"{Context.TmpFolder}\\{imageId}.png";
            return imagePath;
        }
        internal static bool allNulls(object obj1, object obj2, object obj3 = null, object obj4 = null, object obj5 = null, object obj6 = null)
        {
            return (obj1 == null && obj2 == null && obj3 == null && obj4 == null && obj5 == null && obj6 == null);

        }
        internal EPPlusHelper GetHelperForExistingFile(string filePath, FileTypes fileType)
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
        internal void ThrowExpetionsForMissingWorksheet(EPPlusHelper ePPlusHelper, string worksheetName, FileTypes fileType)
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
        internal void ThrowExpetionsForMissingHeader(EPPlusHelper ePPlusHelper, string worksheetName, FileTypes fileType, int rowWithHeaders, List<string> expectedColumns)
        {
            var columnsList = ePPlusHelper.GetHeaders(worksheetName, rowWithHeaders);
            foreach (var expectedColumn in expectedColumns)
            {
                if (!columnsList.Any(_ => _.Equals(expectedColumn, StringComparison.InvariantCultureIgnoreCase)))
                {
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
                            errorType: ErrorTypes.MissingValue,
                            userMessage: $"The file '{fileType}' does not have one of the expected headers ('{expectedColumn}')"
                            );
                }
            }
        }
        #endregion
    }
}