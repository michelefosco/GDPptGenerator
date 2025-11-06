using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using ExcelImageExtractors.Interfaces;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;


namespace ExcelImageExtractors
{
    public class ImageExtractor_Aspose : IImageExtractor
    {
        Workbook workbook;

        public ImageExtractor_Aspose(string excelFilePath)
        {
            // Carica il workbook
            workbook = new Workbook(excelFilePath);
        }

        public void TryToExportToImageFileOnFileSystem(string workSheetName, string printArea, string destinationPath)
        {
            // Seleziono il foglio
            var worksheet = workbook.Worksheets[workSheetName];

            // Seleziono la PrintArea, ovvero la porzione di foglio da esportare
            worksheet.PageSetup.PrintArea = printArea;

            // Opzioni di rendering in immagine
            var imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                OnePagePerSheet = true,
                PrintingPage = PrintingPageType.Default
            };

            // Crea un oggetto SheetRender per il foglio
            var sr = new SheetRender(worksheet, imgOptions);

            // Esporta la pivot (tutto il foglio) come immagine
            var toBeChoppedImagePath = Path.Combine(Path.GetDirectoryName(destinationPath), Path.GetFileNameWithoutExtension(destinationPath) + "_toBeChopped" + Path.GetExtension(destinationPath));
            sr.ToImage(0, toBeChoppedImagePath);

            // Ritaglio l'immagine per eliminare
            chopImage(toBeChoppedImagePath, destinationPath);

            File.Delete(toBeChoppedImagePath);
        }

        public void Close()
        {
            workbook.Dispose();
        }

        private void chopImage(string inputPath, string outputPath)
        {
            using (Bitmap original = new Bitmap(inputPath))
            {
                // Definisci i margini da tagliare
                int top = 72;
                int left = 68;
                int right = 67;
                int bottom = 72;

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
