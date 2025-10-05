using FilesEditor.Enums;
using System.Collections.Generic;

namespace FilesEditor.Entities.MethodsArgs
{
    public class ValidaSourceFilesOutput : UserInterfaceIOutputBase
    {
        public ValidaSourceFilesOutput(EsitiFinali esito) : base(esito)
        { }

        public UserOptions UserOptions;

    }
}