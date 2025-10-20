using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;


namespace ImagesFromExcelGenerator
{
    public class ImageExporter
    {
        Workbook workbook;

        public ImageExporter(string excelFilePath)
        {
            // Carica il workbook
            workbook = new Aspose.Cells.Workbook(excelFilePath);
        }

        public void ExportImages(string workSheetName, string printArea, string destinationPath)
        {
            var worksheet = workbook.Worksheets[workSheetName];

            // per il tipo 5 (chart a parte) imposto un'area di stampa più grande
            worksheet.PageSetup.PrintArea = printArea;


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
            var toBeChoppedImagePath = Path.Combine(Path.GetDirectoryName(destinationPath), Path.GetFileNameWithoutExtension(destinationPath) + "_toBeChopped" + Path.GetExtension(destinationPath));
            sr.ToImage(0, toBeChoppedImagePath);

            chopImage(toBeChoppedImagePath, destinationPath);

            //todo uncommentare
            //File.Delete(toBeChoppedImagePath);
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
