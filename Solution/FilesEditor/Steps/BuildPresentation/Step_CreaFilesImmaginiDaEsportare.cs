using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using FilesEditor.Entities;
using FilesEditor.Enums;
using System;
using System.Drawing;
using System.Drawing.Imaging;

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

            // Carica il workbook
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(Context.DataSourceFilePath);

            foreach (var itemsToExportAsImage in Context.ItemsToExportAsImage)
            {
                var worksheet = workbook.Worksheets[itemsToExportAsImage.Sheet];

                // per il tipo 5 (chart a parte) imposto un'area di stampa più grande
                worksheet.PageSetup.PrintArea = itemsToExportAsImage.PrintArea;


                // Opzioni di rendering in immagine
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    OnePagePerSheet = true,
                    PrintingPage = PrintingPageType.Default
                };
                // Crea un oggetto SheetRender per il foglio
                SheetRender sr = new SheetRender(worksheet, imgOptions);
                //   sr = new SheetRender(worksheetWithPivot, imgOptions);

                // Esporta la pivot (tutto il foglio) come immagine
                var toBeChoppedImagePath = GetImagePath(itemsToExportAsImage.ImageId + "_toBeChopped");
                sr.ToImage(0, toBeChoppedImagePath);

                var finalImagePath = GetImagePath(itemsToExportAsImage.ImageId);
                chopImage(toBeChoppedImagePath, finalImagePath);

                //todo uncommentare
                //File.Delete(toBeChoppedImagePath);
            }
        }


        private void chopImage(string inputPath, string outputPath)
        {
            using (Bitmap original = new Bitmap(inputPath))
            {
                // Definisci i margini da tagliare
                int top = 70;
                int left = 66;
                int right = 65;
                int bottom = 68;

                // Calcola la nuova area utile
                int newWidth = original.Width - left - right;
                int newHeight = original.Height - top - bottom;

                if (newWidth <= 0 || newHeight <= 0)
                {
                    throw new InvalidOperationException("I margini da tagliare sono troppo grandi rispetto all'immagine.");
                }

                Rectangle cropArea = new Rectangle(left, top, newWidth, newHeight);

                // Clona la porzione scelta
                using (Bitmap cropped = original.Clone(cropArea, original.PixelFormat))
                {
                    cropped.Save(outputPath, ImageFormat.Png);
                }
            }
        }
    }
}