using FilesEditor.Entities;
using FilesEditor.Enums;
using System;

namespace FilesEditor.Steps
{
    abstract public class StepBase
    {
        internal StepContext Context;

        internal StepBase(StepContext context)
        {
            Context = context;
        }

        internal abstract string StepName { get; }
        internal abstract EsitiFinali DoSpecificStepTask();

        internal virtual void BeforeTask() { }

        internal virtual void AfterTask() { }


        internal EsitiFinali DoStepTask()
        {
            // Log della informazini di contesto prima dell'esecuzione del task
            Context.DebugInfoLogger.LogStepContext(StepName, Context);

            // Operazioni da farsi prima dell'esecuzione del task (esempio log delle info prima dell'esecuzione del task)
            BeforeTask();

            // Monitoro il tempo impiegato ad eseguire il task
            var startTime = DateTime.UtcNow;
            var result = DoSpecificStepTask();

            // Calcolo e loggo il tempo impiegato per eseguire il task
            var timeSpent = DateTime.UtcNow - startTime;
            Context.DebugInfoLogger.LogPerformance(StepName, timeSpent);

            // Operazioni da farsi dopo l'esecuzione del task (esempio log delle info dopo l'esecuzione del task)
            AfterTask();

            // Log della informazini di contesto dopo l'esecuzione del task
            Context.DebugInfoLogger.LogStepContext(StepName, Context);

            return result;
        }

        #region Utilities
        internal string GetTmpFolderImagePathByImageId(string tmpFolderPath, string imageId)
        {
            var imagePath = $"{tmpFolderPath}\\{imageId}.png";
            return imagePath;
        }
        #endregion
    }
}