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
        internal override string StepName => "Step_DataSource_Editing_Start";

        public Step_DataSource_Editing_Start(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            Context.DataSourceEPPlusHelper.ExcelPackage.Workbook.CalcMode = OfficeOpenXml.ExcelCalcMode.Manual;

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }
    }
}