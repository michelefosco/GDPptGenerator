using FilesEditor.Entities;
using FilesEditor.Enums;
using System;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_ImpostaVarabiliInNameManager : StepBase
    {
        public override string StepName => "Step_ImpostaVarabiliInNameManager";

        internal override void BeforeTask()
        {
            Context.DebugInfoLogger.LogStepContext(StepName, Context);
        }

        internal override void ManageInfoAboutPerformedStepTask(TimeSpan timeSpent)
        {
            Context.DebugInfoLogger.LogPerformance(StepName, timeSpent);
        }

        internal override void AfterTask()
        {
            Context.DebugInfoLogger.LogStepContext(StepName, Context);
        }
        public Step_ImpostaVarabiliInNameManager(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            const string VARIABLE_NAME_ANNO = "anno";
            const string VARIABLE_NAME_MESE = "mese";
            const string VARIABLE_NAME_QUARTER = "quarter";

            Context.DataSourceEPPlusHelper.SetVariableInNameManager(VARIABLE_NAME_ANNO, Context.PeriodYear.ToString());
            Context.DataSourceEPPlusHelper.SetVariableInNameManager(VARIABLE_NAME_MESE, Context.PeriodMont.ToString());
            Context.DataSourceEPPlusHelper.SetVariableInNameManager(VARIABLE_NAME_QUARTER, Context.PeriodQuarter.ToString());

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }
    }
}