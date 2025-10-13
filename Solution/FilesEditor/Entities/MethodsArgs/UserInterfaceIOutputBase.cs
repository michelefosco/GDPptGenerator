using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using FilesEditor.Steps.BuildPresentation;


namespace FilesEditor.Entities.MethodsArgs
{
    public class UserInterfaceIOutputBase
    {
        //public UserInterfaceIOutputBase() { }

        public UserInterfaceIOutputBase(EsitiFinali esito)
        {
            Esito = esito;
        }

        public UserInterfaceIOutputBase(ManagedException managedException)
        {
            Esito = EsitiFinali.Failure;
            ManagedException = managedException;
        }

        public ManagedException ManagedException { get; private set; }
        public EsitiFinali Esito { get; private set; }


        //public void SettaManagedException(ManagedException managedException)
        //{
        //    ManagedException = managedException;
        //    if (managedException != null)
        //    { Esito = EsitiFinali.Failure; }
        //}

        public void SettaEsitoFinale(EsitiFinali esito)
        {
            Esito = esito;
        }
    }
}