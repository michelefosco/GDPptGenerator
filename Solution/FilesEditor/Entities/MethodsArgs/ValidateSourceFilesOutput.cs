using FilesEditor.Entities.Exceptions;
using System.Collections.Generic;

namespace FilesEditor.Entities.MethodsArgs
{
    public class ValidateSourceFilesOutput : UserInterfaceOutputBase
    {
        public List<InputDataFilters_Item> Applicablefilters;

        public ValidateSourceFilesOutput(StepContext context) : base(context)
        {
            Applicablefilters = context.Applicablefilters;
        }

        public ValidateSourceFilesOutput(ManagedException managedException) : base(managedException)
        {
        }
    }
}