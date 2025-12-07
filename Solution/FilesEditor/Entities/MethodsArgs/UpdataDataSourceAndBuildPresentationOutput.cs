using FilesEditor.Entities.Exceptions;
using System.Collections.Generic;

namespace FilesEditor.Entities.MethodsArgs
{
    public class UpdataDataSourceAndBuildPresentationOutput : UserInterfaceOutputBase
    {
        public List<string> OutputFilePathLists = new List<string>();

        internal UpdataDataSourceAndBuildPresentationOutput(StepContext context) : base(context)
        {
            OutputFilePathLists = context.OutputFilePathLists;
        }

        internal UpdataDataSourceAndBuildPresentationOutput(ManagedException managedException) : base(managedException)
        {
        }
    }
}