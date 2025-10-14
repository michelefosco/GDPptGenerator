using FilesEditor.Entities.Exceptions;
using System.Collections.Generic;

namespace FilesEditor.Entities.MethodsArgs
{
    public class ValidaSourceFilesOutput : UserInterfaceOutputBase
    {
        public ValidaSourceFilesOutput(StepContext context) : base(context.Esito)
        {
            Applicablefilters = context.Applicablefilters;
        }

        public ValidaSourceFilesOutput(ManagedException managedException) : base(managedException)
        { }

        //public UserOptions UserOptions;
        public List<InputDataFilters_Items> Applicablefilters;
    }
}