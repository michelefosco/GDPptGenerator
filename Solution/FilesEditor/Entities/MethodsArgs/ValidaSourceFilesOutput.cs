using FilesEditor.Entities.Exceptions;

namespace FilesEditor.Entities.MethodsArgs
{
    public class ValidaSourceFilesOutput : UserInterfaceOutputBase
    {
        public ValidaSourceFilesOutput(StepContext context) : base(context.Esito)
        {
            UserOptions = context.UserOptions;
        }

        public ValidaSourceFilesOutput(ManagedException managedException) : base(managedException)
        { }

        public UserOptions UserOptions;
    }
}