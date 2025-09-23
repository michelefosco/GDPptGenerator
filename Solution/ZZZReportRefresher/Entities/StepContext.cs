using ReportRefresher.Helpers;
using System.Collections.Generic;

namespace ReportRefresher.Entities
{
    internal class StepContext
    {
        public readonly Configurazione Configurazione;
        //
        public readonly UpdateReportsInput UpdateReportsInput;
        public readonly UpdateReportsOutput UpdateReportsOutput;    // oggetto di output, qui vengono messe tutte le informazioni di output utili per interfaccia e controlli dei test
        //
        public FileDebugHelper DebugInfoLogger = new FileDebugHelper(null);
        public InfoFileController InfoFileController;
        public InfoFileReport InfoFileReport;
        //
        public Dictionary<string, object> Parameters = new Dictionary<string, object>();
        //
        public List<string> CategorieFornitori;
        //
        public List<Reparto> RepartiCensitiInController;
        public List<Reparto> RepartiCensitiInReport;
        //
        public List<FornitoreCensito> FornitoriCensitiInReport;
        public List<FornitoreNonCensito> FornitoriNonCensitiInReport;
        //
        public List<RigaSpese> RigheSpese;
        public List<RigaSpeseSkippata> RigheSpeseSkippate;
        //
        public List<RigaTabellaAvanzamento> RigheTabellaAvanzamento;
        public List<RigaTabellaConsumiSpacchettati> RigheTabellaConsumiSpacchettati;
        //
        public StepContext(UpdateReportsInput updateReportsInput, Configurazione configurazione)
        {
            Configurazione = configurazione;

            UpdateReportsInput = updateReportsInput;

            UpdateReportsOutput = new UpdateReportsOutput();
            UpdateReportsOutput.SettaConfigurazioneUsata(Configurazione);
        }
    }
}