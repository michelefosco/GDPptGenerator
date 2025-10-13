using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Entities.MethodsArgs;

namespace FilesEditor.Steps.BuildPresentation 
{
    //todo: rendere internal ma accessibile ai test
    abstract public class BuildPresentation_StepBase : StepBase
    {
        internal BuildPresentation_StepBase(StepContext context)
        {
         base.Context = context;
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
    }
}