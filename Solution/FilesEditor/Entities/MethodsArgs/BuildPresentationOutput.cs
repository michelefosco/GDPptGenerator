using FilesEditor.Entities.Exceptions;
using System.Collections.Generic;

namespace FilesEditor.Entities.MethodsArgs
{
    public class BuildPresentationOutput : UserInterfaceOutputBase
    {
        public List<string> OutputFilePathLists = new List<string>();

        internal BuildPresentationOutput(StepContext context) : base(context)
        {
            OutputFilePathLists = context.OutputFilePathLists;
        }

        internal BuildPresentationOutput(ManagedException managedException) : base(managedException)
        {
        }
    }
}