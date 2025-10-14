using FilesEditor.Entities;
using FilesEditor.Enums;

namespace FilesEditor.Steps
{
    /// <summary>
    /// Step che se raggiunto indica l'esito Success dell'intero processo
    /// </summary>
    internal class Step_EsitoFinale_Success : StepBase
    {
        public Step_EsitoFinale_Success(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            return EsitiFinali.Success;
        }
    }
}