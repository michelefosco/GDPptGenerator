using FilesEditor.Entities;
using FilesEditor.Enums;
using System;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_DataSource_Editing_Stop : StepBase
    {
        internal override string StepName => "Step_DataSource_Editing_Stop";

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