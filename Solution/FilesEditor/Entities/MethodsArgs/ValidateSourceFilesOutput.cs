using FilesEditor.Entities.Exceptions;
using System.Collections.Generic;

namespace FilesEditor.Entities.MethodsArgs
{
    public class ValidateSourceFilesOutput : UserInterfaceOutputBase
    {
        public List<InputDataFilters_Item> Applicablefilters;

        internal ValidateSourceFilesOutput(StepContext context, ManagedException managedException= null) : base(context)
        {
            Applicablefilters = new List<InputDataFilters_Item>(context.ApplicableFilters);
        }
    }
}