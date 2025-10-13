using FilesEditor.Entities;
using FilesEditor.Enums;
using FilesEditor.Entities.MethodsArgs;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// Step che se raggiungo indica l'esito Success dell'intero processo
    /// </summary>
    internal class Step_EsitoFinaleSuccess : Step_Base
    {
        public Step_EsitoFinaleSuccess(StepContext context) : base(context)
        { }

        internal override BuildPresentationOutput DoSpecificTask()
        {
            return FinalizzaOutput(EsitiFinali.Success);
        }
    }
}