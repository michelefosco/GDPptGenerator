using FilesEditor.Entities;
using FilesEditor.Enums;

namespace FilesEditor.Steps
{
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
        internal string GetTmpFolderImagePathByImageId(string tmpFolderPath, string imageId)
        {
            var imagePath = $"{tmpFolderPath}\\{imageId}.png";
            return imagePath;
        }
        #endregion
    }
}