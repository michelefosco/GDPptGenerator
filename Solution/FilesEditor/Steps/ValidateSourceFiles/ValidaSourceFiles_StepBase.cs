using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Entities.MethodsArgs;

namespace FilesEditor.Steps.ValidateSourceFiles
{
    abstract public class ValidateSourceFiles_StepBase : StepBase
    {
        internal ValidateSourceFiles_StepBase(StepContext context)
        {
            base.Context = context;
        }

        internal abstract ValidaSourceFilesOutput DoSpecificTask();

        internal ValidaSourceFilesOutput Do()
        {
            try
            {
                return DoSpecificTask();
            }
            catch (ManagedException managedException)
            {
                return new ValidaSourceFilesOutput(managedException);
            }
        }
    }
}