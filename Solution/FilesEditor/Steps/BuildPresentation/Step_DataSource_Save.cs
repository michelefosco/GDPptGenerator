using FilesEditor.Entities;
using FilesEditor.Enums;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_DataSource_Save : StepBase
    {
        internal override string StepName => "Step_DataSource_Save";

        public Step_DataSource_Save(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            Context.DataSourceEPPlusHelper.ExcelPackage.Workbook.CalcMode = OfficeOpenXml.ExcelCalcMode.Automatic;
            Context.DataSourceEPPlusHelper.ExcelPackage.Workbook.FullCalcOnLoad = true; // true è comunque il default

            Context.DataSourceEPPlusHelper.Save();
            Context.DataSourceEPPlusHelper.Close();

            Context.SetDatasourceStatus_ImportDatiCompletato();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }
    }
}