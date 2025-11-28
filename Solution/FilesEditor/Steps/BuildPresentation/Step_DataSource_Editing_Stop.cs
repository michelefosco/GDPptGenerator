using FilesEditor.Entities;
using FilesEditor.Enums;
using System;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_DataSource_Editing_Stop: StepBase
    {
        public override string StepName => "Step_DataSource_Editing_Stop";

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

        public Step_DataSource_Editing_Stop(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            Context.DataSourceEPPlusHelper.ExcelPackage.Workbook.CalcMode = OfficeOpenXml.ExcelCalcMode.Automatic;
            Context.DataSourceEPPlusHelper.ExcelPackage.Workbook.FullCalcOnLoad = true; // true è comunque il default

            Context.DataSourceEPPlusHelper.Save();
            Context.DataSourceEPPlusHelper.Close();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }
    }
}