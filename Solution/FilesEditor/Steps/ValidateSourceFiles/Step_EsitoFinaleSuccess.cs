using FilesEditor.Entities;
using FilesEditor.Enums;
using FilesEditor.Entities.MethodsArgs;

namespace FilesEditor.Steps.ValidateSourceFiles
{
    /// <summary>
    /// Step che se raggiunto indica l'esito Success dell'intero processo
    /// </summary>
    internal class Step_EsitoFinaleSuccess : ValidateSourceFiles_StepBase
    {
        public Step_EsitoFinaleSuccess(StepContext context) : base(context)
        { }

        internal override ValidaSourceFilesOutput DoSpecificTask()
        {
            Context.SettaEsitoFinale(EsitiFinali.Success);
            return new ValidaSourceFilesOutput(Context);
        }
    }
}