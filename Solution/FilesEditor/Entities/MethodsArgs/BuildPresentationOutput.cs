using FilesEditor.Entities.Exceptions;
using System.Collections.Generic;

namespace FilesEditor.Entities.MethodsArgs
{
    public class BuildPresentationOutput : UserInterfaceOutputBase
    {
        public List<string> OutputFilePathLists = new List<string>();

        public BuildPresentationOutput(StepContext context) : base(context.Esito)
        {
            OutputFilePathLists = context.OutputFilePathLists;
        }

        public BuildPresentationOutput(ManagedException managedException) : base(managedException)
        {
        }
    }
}