
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FilesEditor.Steps.BuildPresentation
{
    //todo: rendere internal ma accessibile ai test
    abstract public class Step_Base
    {
        internal readonly StepContext Context;

        internal Step_Base(StepContext context)
        {
            Context = context;
        }

        internal abstract BuildPresentationOutput DoSpecificTask();

        internal BuildPresentationOutput Do()
        {
            try
            {
                return DoSpecificTask();
            }
            catch (ManagedException managedException)
            {
                return new BuildPresentationOutput(managedException);
                //Context.BuildPresentationOutput.SettaManagedException(ex);
                //return FinalizzaOutput(EsitiFinali.Failure);
            }
        }


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