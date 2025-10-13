using FilesEditor.Entities;
using FilesEditor.Entities.MethodsArgs;
using System.IO;

namespace FilesEditor.Steps.BuildPresentation
{
    internal class Step_AggiornaDataSource : Step_Base
    {
        public Step_AggiornaDataSource(StepContext context) : base(context)
        { }

        internal override BuildPresentationOutput DoSpecificTask()
        {
            aggiornaDataSource();
            return null; // Step intermedio, non ritorna alcun esito
        }

        private void aggiornaDataSource()
        {
            //todo: inserire logica di backtup del file esistente
            // Simulazione del creazione del file Excel con i dati...
            Context.OutputDataSourceFilePath = Path.Combine(Context.TmpFolder, Constants.FileNames.DATA_SOURCE_TEMPLATE_FILENAME);
            var dataSourceTemplateFile = Path.Combine(Context.SourceFilesFolder, Constants.FileNames.DATA_SOURCE_TEMPLATE_FILENAME);
            File.Copy(dataSourceTemplateFile, Context.OutputDataSourceFilePath, true);
        }
    }
}