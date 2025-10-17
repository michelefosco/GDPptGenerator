using FilesEditor.Entities;
using FilesEditor.Enums;
using ImagesFromExcelGenerator;

namespace FilesEditor.Steps.BuildPresentation
{
    internal class Step_CreaFilesImmaginiDaEsportare : StepBase
    {
        public Step_CreaFilesImmaginiDaEsportare(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            Context.DebugInfoLogger.LogStepContext("Step_CreaFilesImmaginiDaEsportare", Context);
            creaFilesImmaginiDaEsportare();
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        /// <summary>
        /// Genera le immagini da usarsi nelle slides
        /// </summary>
        private void creaFilesImmaginiDaEsportare()
        {
            Exporter imgExporter = new Exporter(Context.DataSourceFilePath);
            foreach (var itemsToExportAsImage in Context.ItemsToExportAsImage)
            {
                imgExporter.ExportImages(itemsToExportAsImage.Sheet, itemsToExportAsImage.PrintArea, GetImagePath(itemsToExportAsImage.ImageId));
            }
        }
    }
}