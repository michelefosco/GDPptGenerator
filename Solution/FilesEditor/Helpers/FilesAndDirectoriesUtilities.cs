using FilesEditor.Constants;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FilesEditor.Helpers
{
    internal class FilesAndDirectoriesUtilities
    {
        static internal void CancellaFileSeEsiste(string filePath, FileTypes fileType)
        {
            if (File.Exists(filePath))
            {
                try
                {
                    File.Delete(filePath);
                }
                catch (Exception)
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

        static internal void CancellaDirectorySeEsiste(string folderPath)
        {
            if (Directory.Exists(folderPath))
            {
                try
                {
                    Directory.Delete(folderPath, true);
                }
                catch (Exception)
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
                        errorType: ErrorTypes.UnableToDeleteFolder,
                        userMessage: string.Format(UserErrorMessages.UnableToDeleteFolder, folderPath)
                        );
                }
            }
        }

        static internal void CreaDirectorySeNonEsiste(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                try
                {
                    Directory.CreateDirectory(folderPath);
                }
                catch (Exception)
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
                        userMessage: string.Format(UserErrorMessages.UnableToCreateFolder, folderPath)
                        );
                }
            }
        }

        /// <summary>
        /// Get all files matching the criteria
        /// </summary>
        /// <param name="folderPath"></param>
        /// <param name="filter"></param>
        /// <returns></returns>
        static internal List<string> GetFilesListFromFolder(string folderPath, string filter)
        {
            // rimuovo dalla lista i file il cui nome inizia con "~$" (ovvero i file temporaranei creati da Excel quando un file è aperto)
            var filePaths = Directory.GetFiles(folderPath, filter, SearchOption.TopDirectoryOnly)
                    .Where(_ => !Path.GetFileName(_).StartsWith("~$")).ToList();
            return filePaths;
        }
    }
}
