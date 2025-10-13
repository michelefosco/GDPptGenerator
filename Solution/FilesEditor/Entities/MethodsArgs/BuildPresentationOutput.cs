using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;

namespace FilesEditor.Entities.MethodsArgs
{
    public class BuildPresentationOutput : UserInterfaceIOutputBase
    {
        public BuildPresentationOutput(EsitiFinali esito) : base(esito)
        { }

        public BuildPresentationOutput(ManagedException managedException) : base(managedException)
        { }

        //public List<Reparto> RepartiCensitiInController { get; private set; }
        //public List<Reparto> RepartiCensitiInReport { get; private set; }
        //public List<FornitoreCensito> FornitoriCensitiInReport { get; private set; }
        //public List<FornitoreNonCensito> FornitoriNonCensitiInReport { get; private set; }
        //public List<RigaSpese> RigheSpesa { get; private set; }
        //  public List<RigaSpeseSkippata> RigheSpesaSkippate { get; private set; }
        //public List<RigaTabellaAvanzamento> RigheTabellaAvanzamento { get; private set; }
        //public List<RigaTabellaConsumiSpacchettati> RigheTabellaConsumiSpacchettati { get; private set; }
        //public List<TuplaSpesePerAnno> RigheTabellaSintesi_SetMinimoRigheNecessarie { get; private set; }
        //public List<RigaTabellaSintesi> RigheTabellaSintesi_RigheMancanti { get; private set; }

        //public List<string> CategorieFornitori { get; private set; }




        //public void SettaRepartiCensitiController(List<Reparto> reparti)
        //{
        //    RepartiCensitiInController = reparti;
        //}

        //public void SettaRepartiCensitiInReport(List<Reparto> reparti)
        //{
        //    RepartiCensitiInReport = reparti;
        //}

        //public void SettaFornitoriCensiti(List<FornitoreCensito> fornitori)
        //{
        //    FornitoriCensitiInReport = fornitori;
        //}

        //public void SettaRigheSpesa(List<RigaSpese> righeSpesa)
        //{
        //    RigheSpesa = righeSpesa;
        //}
        //public void SettaRigheSpesaSkippate(List<RigaSpeseSkippata> righeSpesaSkippate)
        //{
        //    RigheSpesaSkippate = righeSpesaSkippate;
        //}

        //public void SettaTabellaAvanzamento(List<RigaTabellaAvanzamento> righeTabellaAvanzamento)
        //{
        //    RigheTabellaAvanzamento = righeTabellaAvanzamento;
        //}

        //public void SettaTabellaConsumiSpacchettati(List<RigaTabellaConsumiSpacchettati> righeTabellaConsumiSpacchettati)
        //{
        //    RigheTabellaConsumiSpacchettati = righeTabellaConsumiSpacchettati;
        //}

        //public void SettaTabellaSintesi_SetMinimoRigheNecessarie(List<TuplaSpesePerAnno> righeTabellaSintesi)
        //{
        //    RigheTabellaSintesi_SetMinimoRigheNecessarie = righeTabellaSintesi;
        //}

        //public void SettaTabellaSintesi_RigheMancanti(List<RigaTabellaSintesi> righeTabellaSintesi)
        //{
        //    RigheTabellaSintesi_RigheMancanti = righeTabellaSintesi;
        //}

        //public void SettaFornitoriNonCensitiInReport(List<FornitoreNonCensito> fornitoriNonCensitiInReport)
        //{
        //    FornitoriNonCensitiInReport = fornitoriNonCensitiInReport;
        //}

        //public void SettaCategorieFornitori(List<string> categorieFornitori)
        //{
        //    CategorieFornitori = categorieFornitori;
        //}

    }
}