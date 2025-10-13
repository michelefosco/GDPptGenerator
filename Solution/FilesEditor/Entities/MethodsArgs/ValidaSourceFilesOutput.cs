using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using System.Collections.Generic;

namespace FilesEditor.Entities.MethodsArgs
{
    public class ValidaSourceFilesOutput : UserInterfaceOutputBase
    {
        public ValidaSourceFilesOutput(EsitiFinali esito) : base(esito)
        { }

        public ValidaSourceFilesOutput(ManagedException managedException) : base(managedException)
        { }

        public UserOptions UserOptions;
    }
}