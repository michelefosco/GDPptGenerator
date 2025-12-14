using FilesEditor.Entities.Exceptions;
using System.Collections.Generic;

namespace FilesEditor.Entities.MethodsArgs
{
    public class UpdataDataSourceAndBuildPresentationOutput : UserInterfaceOutputBase
    {
        public List<string> OutputFilePathLists = new List<string>();
        public bool DatasourceStatus_ImportDatiCompletato { get; private set; }
        public bool DatasourceStatus_RefreshAllCompletato { get; private set; }

        internal UpdataDataSourceAndBuildPresentationOutput(StepContext context, ManagedException managedException = null) : base(context, managedException)
        {
            OutputFilePathLists = context.OutputFilePathLists;
            DatasourceStatus_ImportDatiCompletato = context.DatasourceStatus_ImportDatiCompletato;
            DatasourceStatus_RefreshAllCompletato = context.DatasourceStatus_RefreshAllCompletato;
        }
    }
}