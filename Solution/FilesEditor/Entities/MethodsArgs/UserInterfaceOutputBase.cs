using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;

namespace FilesEditor.Entities.MethodsArgs
{
    public class UserInterfaceOutputBase
    {
        public UserInterfaceOutputBase(EsitiFinali esito)
        {
            Esito = esito;
        }

        public UserInterfaceOutputBase(ManagedException managedException)
        {
            Esito = EsitiFinali.Failure;
            ManagedException = managedException;
        }

        public ManagedException ManagedException { get; private set; }
        public EsitiFinali Esito { get; private set; }
    }
}