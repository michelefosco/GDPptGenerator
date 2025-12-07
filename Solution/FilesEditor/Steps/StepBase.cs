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

        public abstract string StepName { get; }

        internal abstract void BeforeTask();
        internal abstract EsitiFinali DoSpecificStepTask();
        internal abstract void ManageInfoAboutPerformedStepTask(TimeSpan timeSpent);

        internal abstract void AfterTask();

        internal EsitiFinali DoStepTask()
        {
            // Operazioni da farsi prima dell'esecuzione del task (esempio log delle info prima dell'esecuzione del task)
            BeforeTask();

            // Monitoro il tempo impiegato ad eseguire il task
            var startTime = DateTime.UtcNow;
            var result = DoSpecificStepTask();
            var endTime = DateTime.UtcNow;

            // Passo le info sul tempo impiegato al task
            var timeSpent = endTime - startTime;
            ManageInfoAboutPerformedStepTask(timeSpent);

            // Operazioni da farsi dopo l'esecuzione del task (esempio log delle info dopo l'esecuzione del task)
            AfterTask();

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