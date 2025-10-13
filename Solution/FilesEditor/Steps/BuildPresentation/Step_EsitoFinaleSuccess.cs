using FilesEditor.Entities;
using FilesEditor.Enums;
using FilesEditor.Entities.MethodsArgs;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// Step che se raggiungo indica l'esito Success dell'intero processo
    /// </summary>
    internal class Step_EsitoFinaleSuccess : BuildPresentation_StepBase
    {
        public Step_EsitoFinaleSuccess(StepContext context) : base(context)
        { }

        internal override BuildPresentationOutput DoSpecificTask()
        {
            Context.SettaEsitoFinale(EsitiFinali.Success);
            return new BuildPresentationOutput(Context);
        }
    }
}