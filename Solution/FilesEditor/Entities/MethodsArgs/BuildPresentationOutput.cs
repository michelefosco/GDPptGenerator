using FilesEditor.Entities.Exceptions;

namespace FilesEditor.Entities.MethodsArgs
{
    public class BuildPresentationOutput : UserInterfaceOutputBase
    {
        public BuildPresentationOutput(StepContext context) : base(context.Esito)
        {
        }

        public BuildPresentationOutput(ManagedException managedException) : base(managedException)
        {
        }
    }
}