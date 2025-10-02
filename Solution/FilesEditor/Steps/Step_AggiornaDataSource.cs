using FilesEditor.Entities;
using System.IO;

namespace FilesEditor.Steps
{
    internal class Step_AggiornaDataSource : Step_Base
    {
        public Step_AggiornaDataSource(StepContext context) : base(context)
        { }

        internal override CreatePresentationsOutput DoSpecificTask()
        {
            aggiornaDataSource();
            return null; // Step intermedio, non ritorna alcun esito
        }

        private void aggiornaDataSource()
        {
            // Simulazione del creazione del file Excel con i dati...
            Context.ExcelDataSourceFile = Path.Combine(Context.TmpFolder, Constants.FileNames.DATA_SOURCE_FILENAME);
            var dataSourceTemplateFile = Path.Combine(Context.TemplatesFolder, Constants.FileNames.DATA_SOURCE_FILENAME);
            File.Copy(dataSourceTemplateFile, Context.ExcelDataSourceFile, true);
        }
    }
}