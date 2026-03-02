using FilesEditor.Entities;
using FilesEditor.Enums;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_ImpostaVarabiliInNameManager : StepBase
    {
        internal override string StepName => "Step_ImpostaVarabiliInNameManager";

        public Step_ImpostaVarabiliInNameManager(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            const string VARIABLE_NAME_ANNO = "anno";
            const string VARIABLE_NAME_MESE = "mese";
            const string VARIABLE_NAME_QUARTER = "quarter";

            Context.DataSourceEPPlusHelper.SetVariableInNameManager(VARIABLE_NAME_ANNO, Context.PeriodYear);
            Context.DataSourceEPPlusHelper.SetVariableInNameManager(VARIABLE_NAME_MESE, Context.PeriodMont);
            Context.DataSourceEPPlusHelper.SetVariableInNameManager(VARIABLE_NAME_QUARTER, Context.PeriodQuarter);

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }
    }
}