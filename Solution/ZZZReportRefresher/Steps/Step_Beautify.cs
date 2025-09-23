using ReportRefresher.Entities;

namespace ReportRefresher.Steps
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_Beautify : Step_Base
    {
        internal override UpdateReportsOutput DoSpecificTask(StepContext context)
        {
            Beautify(context.InfoFileReport);

            return null;
        }

        private void Beautify(InfoFileReport infoFileReport)
        {
            foreach (var worksheetName in infoFileReport.EPPlusHelper.GetWorksheetNames())
            {
                // Porta il cursore della cella selezionata sulla prima riga e colonna di ogni foglio
                infoFileReport.EPPlusHelper.SelectWorksheet(worksheetName, 1, 1);
            }

            // Selezione il foglio "Anagrafica fornitori"
            infoFileReport.EPPlusHelper.SelectWorksheet(infoFileReport.WorksheetName_AnagraficaFornitori);
        }
    }
}