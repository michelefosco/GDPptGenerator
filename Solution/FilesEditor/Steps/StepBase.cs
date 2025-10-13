using FilesEditor.Entities;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FilesEditor.Steps
{
    //todo: rendere internal ma accessibile ai test
    abstract public class StepBase
    {
        internal StepContext Context;

        internal void AddWarning(string warningMessage)
        {
            //Context.UpdateReportsOutput.Warnings.Add(warningMessage);
            //Context.FileDebugHelper?.LogWarning(warningMessage);
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
            var imagePath = $"{Context.TmpFolder}\\{imageId}.png";
            return imagePath;
        }
    }
}