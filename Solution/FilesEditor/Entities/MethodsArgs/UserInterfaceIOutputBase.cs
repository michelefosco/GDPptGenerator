using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;


namespace FilesEditor.Entities.MethodsArgs
{
    public class UserInterfaceIOutputBase
    {
        public ManagedException ManagedException { get; private set; }
        public EsitiFinali Esito { get; private set; }


        public void SettaManagedException(ManagedException managedException)
        {
            ManagedException = managedException;
            if (managedException != null)
            { Esito = EsitiFinali.Failure; }
        }

        public void SettaEsitoFinale(EsitiFinali esito)
        {
            Esito = esito;
        }
    }
}