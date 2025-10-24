using FilesEditor.Entities;
using FilesEditor.Enums;

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
            // Aspose.Cell
            //var imgExporter = new ExcelImageExtractor.ImageExtractor(Context.DataSourceFilePath);

            var imgExporter = new ExcelImageExtractorInterOp.ExcelImageSaver(Context.DataSourceFilePath);

            foreach (var itemsToExportAsImage in Context.ItemsToExportAsImage)
            {
                var imagePath = GetTmpFolderImagePathByImageId(Context.TmpFolder, itemsToExportAsImage.ImageId);
                imgExporter.ExportImages(itemsToExportAsImage.Sheet, itemsToExportAsImage.PrintArea, imagePath);
            }
            imgExporter.Close();
        }
    }
}