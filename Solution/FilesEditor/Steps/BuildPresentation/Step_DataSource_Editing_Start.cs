using FilesEditor.Entities;
using FilesEditor.Enums;
using System;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_DataSource_Editing_Start : StepBase
    {
        public override string StepName => "Step_DataSource_Editing_Start";

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

        public Step_DataSource_Editing_Start(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            Context.DataSourceEPPlusHelper.ExcelPackage.Workbook.CalcMode = OfficeOpenXml.ExcelCalcMode.Manual;

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }
    }
}