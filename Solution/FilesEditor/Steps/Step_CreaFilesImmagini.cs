using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using FilesEditor.Entities;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace FilesEditor.Steps
{
    internal class Step_CreaFilesImmagini : Step_Base
    {
        public Step_CreaFilesImmagini(StepContext context) : base(context)
        { }

        internal override CreatePresentationsOutput DoSpecificTask()
        {
            creaFilesImmagini();
            return null; // Step intermedio, non ritorna alcun esito
        }

        const string OUTPUT_POWERPOINT_FILENAME = "Output.pptx";

        private void creaFilesImmagini()
        {
            // Lettura del file di testo con le aree di stampa e creazione della lista delle immagini da generare
            creaListaImmaginiDaGenerare();

            // Genera le immagini da usarsi nelle slides
            generaImmagini();
        }


        private void creaListaImmaginiDaGenerare()
        {
            string percorsoFile = Path.Combine(Context.ConfigurationFolder, Constants.FileNames.DATASOURCE_PRINT_AREAS_FILENAME);

            if (File.Exists(percorsoFile))
            {
                Context.ItemsToExportAsImage = new List<ItemToExport>();

                // Legge tutte le righe del file
                string[] righe = File.ReadAllLines(percorsoFile);

                foreach (string riga in righe)
                {
                    // Divide la riga nei campi separati da ";"
                    string[] campi = riga.Split(';');


                    //todo: ragionare su queste trasformazioni
                    var imageId = campi[0].Trim().ToLower();
                    var sheet = campi[1].Trim();
                    var printArea = campi[2].Trim().ToUpper();

                    Context.ItemsToExportAsImage.Add(new ItemToExport { ImageId = imageId, Sheet = sheet, PrintArea = printArea, });
                }
            }
            else
            {
                //todo: eccezione managed
                Console.WriteLine("Il file non esiste!");
            }
        }

        //private void GeneraimmaginiPivotTables()
        //{
        //    // Carica il workbook
        //    Workbook workbook = new Workbook(Context.ExcelDataSourceFile);

        //    const int PIVOT_TYPES_NUMBER = 4;

        //    // context.GeneratedImagesList = new List<GeneratedImages>();

        //    for (int pivotType = 1; pivotType <= PIVOT_TYPES_NUMBER; pivotType++)
        //    {
        //        var worksheetName = $"Shapes_{pivotType.ToString("D2")}";
        //        var worksheetWithPivot = workbook.Worksheets[worksheetName];

        //        // Prendi la prima PivotTable del foglio
        //        PivotTable pivot = worksheetWithPivot.PivotTables[0];

        //        // Ottieni il range della pivot
        //        CellArea range = pivot.TableRange2;

        //        // Imposta l'area di stampa sul range della pivot
        //        worksheetWithPivot.PageSetup.PrintArea =
        //            CellsHelper.CellIndexToName(range.StartRow, range.StartColumn) + ":" +
        //            CellsHelper.CellIndexToName(range.EndRow, range.EndColumn);

        //        // Opzioni di rendering in immagine
        //        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        //        {
        //            ImageType = ImageType.Png,
        //            OnePagePerSheet = true,
        //            PrintingPage = PrintingPageType.Default
        //        };
        //        // Crea un oggetto SheetRender per il foglio
        //        SheetRender sr = new SheetRender(worksheetWithPivot, imgOptions);
        //        //   sr = new SheetRender(worksheetWithPivot, imgOptions);

        //        // Esporta la pivot (tutto il foglio) come immagine
        //        var imagePath = Context.TmpFolder + $"\\Img_{worksheetName}_Pivot.png";
        //        sr.ToImage(0, imagePath);

        //        // aggiungo alla lista delle immagini generate
        //        //context.GeneratedImagesList.Add(new GeneratedImages { Path = imagePath, PivotType = pivotType });
        //    }
        //}


        private void generaImmagini()
        {
            // Carica il workbook
            Workbook workbook = new Workbook(Context.ExcelDataSourceFile);

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
                var toBeChoppedImagePath = getImagePath(itemsToExportAsImage.ImageId + "_toBeChopped");
                sr.ToImage(0, toBeChoppedImagePath);

                var finalImagePath = getImagePath(itemsToExportAsImage.ImageId);
                chopImage(toBeChoppedImagePath, finalImagePath);

                //todo uncommentare
                //File.Delete(toBeChoppedImagePath);
            }
        }

        private string getImagePath(string imageId)
        {
            var imagePath = $"{Context.TmpFolder}\\{imageId}.png";
            return imagePath;
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