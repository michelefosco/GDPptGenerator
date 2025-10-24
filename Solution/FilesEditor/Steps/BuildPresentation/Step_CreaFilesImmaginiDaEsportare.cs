using ExcelImageExtractors.Interfaces;
using FilesEditor.Entities;
using FilesEditor.Enums;
using System.IO;
using System.Linq;

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
            #region Verifico che i file delle immagini da generare per le slide non esistano già
            foreach (var itemsToExportAsImage in Context.ItemsToExportAsImage)
            {
                if (File.Exists(itemsToExportAsImage.ImageFilePath))
                {
                    throw new System.Exception($"The file '{itemsToExportAsImage.ImageFilePath}' should not be there yet!");
                }
            }
            #endregion

            // Predisposta la possibilità di usare Aspose.Cell in casi estremi
            var useIterops = true;
            var imageExtractor = (useIterops)
                ? (IImageExtractor)new ExcelImageExtractors.ImageExtractor(Context.DataSourceFilePath)
                : (IImageExtractor)new ExcelImageExtractors.ImageExtractor_Aspose(Context.DataSourceFilePath);  // Aspose.Cell

            for (int attemptNumber = 1; attemptNumber <= 3; attemptNumber++)
            {
                // Processo tutti gli elementi non ancora presenti sul file system
                foreach (var itemsToExportAsImage in Context.ItemsToExportAsImage.Where(_ => !_.IsPresentOnFileSistem))
                {
                    // tento di generare il file, alcune volte potrebbe non funzionare al primo tentativo
                    imageExtractor.TryToExportToImageFileOnFileSystem(itemsToExportAsImage.WorkSheetName, itemsToExportAsImage.PrintArea, itemsToExportAsImage.ImageFilePath);

                    // verifico la sua presenza su file system ed eventualmente lo marco come "Presente"
                    if (File.Exists(itemsToExportAsImage.ImageFilePath))
                    { itemsToExportAsImage.MarkAsPresentOnFileSistem(); }
                }

                // se tutti i file sono presenti interrompo i tentativi 
                if (!Context.ItemsToExportAsImage.Any(_ => !_.IsPresentOnFileSistem))
                { break; }
            }

            // rilascio le risorse e chiudo il processo di Excel
            imageExtractor.Close();
        }
    }
}