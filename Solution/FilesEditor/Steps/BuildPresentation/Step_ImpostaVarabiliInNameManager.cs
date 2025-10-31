using FilesEditor.Entities;
using FilesEditor.Enums;

namespace FilesEditor.Steps.BuildPresentation
{
    internal class Step_ImpostaVarabiliInNameManager : StepBase
    {
        public Step_ImpostaVarabiliInNameManager(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            Context.DebugInfoLogger.LogStepContext("Step_ImpostaVarabiliInNameManager", Context);

            const string VARIABLE_NAME_ANNO = "anno";
            const string VARIABLE_NAME_MESE = "mese";
            const string VARIABLE_NAME_QUARTER= "quarter";

            Context.EpplusHelperDataSource.SetVariableInNameManager(VARIABLE_NAME_ANNO, Context.PeriodYear.ToString());
            Context.EpplusHelperDataSource.SetVariableInNameManager(VARIABLE_NAME_MESE, Context.PeriodMont.ToString());
            Context.EpplusHelperDataSource.SetVariableInNameManager(VARIABLE_NAME_QUARTER, Context.PeriodQuarter.ToString());
            
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }
    }
}