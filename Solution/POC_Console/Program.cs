using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Pivot;
using Aspose.Cells.Rendering;
using POC_Console.Entities;
//
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;


namespace POC_Console
{
    internal partial class Program
    {

        const int NumberOfTemplateSlides = 4;

        //todo: da mofificare con i percorsi reali, prendere da cartella di esecuzione

        const string PPT_TEMPLATE_FILE_PATH = @"C:\Data\Lefo\Dev\3 GDPptGenerator\Solution\PptGeneratorGUI\bin\Debug\Configurazione\PowerPoint_Template.pptx";
        const string PPT_PRINT_AREAS_FILE_PATH = @"C:\Data\Lefo\Dev\3 GDPptGenerator\Solution\PptGeneratorGUI\bin\Debug\Configurazione\PowerPoint_PrintAreas.txt";
        const string PPT_STRUTTURA_FILE_PATH = @"C:\Data\Lefo\Dev\3 GDPptGenerator\Solution\PptGeneratorGUI\bin\Debug\Configurazione\PowerPoint_Struttura.txt";
        const string XLS_DATASOURCE_FILE_PATH = @"C:\Data\Lefo\Dev\3 GDPptGenerator\Solution\PptGeneratorGUI\bin\Debug\Configurazione\DataSource.xlsx";

        const string OUTPUT_POWERPOINT_FILENAME = "Output.pptx";

        //public class GeneratedImages
        //{
        //    public int PivotType;
        //    public string Path;
        //}

        static Context context;

        static void Main()
        {
            context = new Context();
            // scelta dall'utente...
            context.OutputFolder = @"C:\Data\Lefo\Dev\3 GDPptGenerator\AAA_Demo\AAA Output";

            CreaListaSlidesDaGenerare();

            PredisposiTmpFolder();

            // Creazione della sorgente dati ed aggiornamento degli oggetti correlati (PivotTables, grafici, etc)
            GeneraFileConDati();

            // Lettura del file di testo con le aree di stampa e creazione della lista delle immagini da generare
            CreaListaImmaginiDaGenerare();

            // Genera le immagini da usarsi nelle slides
            GeneraimmaginiItmes();

            // dall'UI
            CreaListaSlidesDaGenerare();

            // Creazione del PPT con le slides scelte dall'utente
            CreazionePowerPoint();
        }

        static void CreaListaSlidesDaGenerare()
        {
            string percorsoFile = PPT_STRUTTURA_FILE_PATH;

            if (File.Exists(percorsoFile))
            {
                context.SlideToGenerateList = new List<SlideToGenerate>();

                // Legge tutte le righe del file
                string[] righe = File.ReadAllLines(percorsoFile);

                foreach (string riga in righe)
                {
                    // Divide la riga nei campi separati da ";"
                    string[] campi = riga.Split(';');

                    //todo: ragionare su queste trasformazioni
                    var slideType = int.Parse(campi[0].Trim());
                    var imageId1 = campi[1].Trim().ToUpper();
                    var imageId2 = campi[2].Trim().ToUpper();
                    var imageId3 = campi[3].Trim().ToLower();


                    context.SlideToGenerateList.Add(new SlideToGenerate
                    {
                        SlideType = slideType,
                        ImageId1 = imageId1,
                        ImageId2 = imageId2,
                        ImageId3 = imageId3
                    });
                }
            }
            else
            {
                //todo: eccezione managed
                Console.WriteLine("Il file non esiste!");
            }

            //return;
            //const int SLIDE_TEMPLATE_1_INDEX = 1;
            //const int SLIDE_TEMPLATE_2_INDEX = 2;
            //const int SLIDE_TEMPLATE_3_INDEX = 3;
            //const int SLIDE_TEMPLATE_4_INDEX = 4;

            //context.SlideToGenerateList = new List<SlideToGenerate>
            //{
            //    new SlideToGenerate { SlideType=SLIDE_TEMPLATE_1_INDEX, ImageId1 = "Img_001" },
            //    new SlideToGenerate { SlideType=SLIDE_TEMPLATE_1_INDEX, ImageId1 = "Img_002" },
            //    new SlideToGenerate { SlideType=SLIDE_TEMPLATE_1_INDEX, ImageId1 = "Img_003" },
            //    new SlideToGenerate { SlideType=SLIDE_TEMPLATE_1_INDEX, ImageId1 = "Img_004" },

            //    new SlideToGenerate { SlideType=SLIDE_TEMPLATE_2_INDEX, ImageId1 = "Img_001",ImageId2 = "Img_002" },
            //    new SlideToGenerate { SlideType=SLIDE_TEMPLATE_2_INDEX, ImageId1 = "Img_003",ImageId2 = "Img_004" },
            //};

            //context.SlideToGenerateList = new List<SlideToGenerate>
            //{
            //    new SlideToGenerate { PivotType = 4, SlideTemplateIndex = SLIDE_TEMPLATE_1_INDEX },
            //    new SlideToGenerate { PivotType = 3, SlideTemplateIndex = SLIDE_TEMPLATE_1_INDEX },
            //    new SlideToGenerate { PivotType = 2, SlideTemplateIndex = SLIDE_TEMPLATE_1_INDEX },
            //    new SlideToGenerate { PivotType = 1, SlideTemplateIndex = SLIDE_TEMPLATE_1_INDEX },

            //    new SlideToGenerate { PivotType = 4, SlideTemplateIndex = SLIDE_TEMPLATE_2_INDEX },
            //    new SlideToGenerate { PivotType = 3, SlideTemplateIndex = SLIDE_TEMPLATE_2_INDEX },
            //    new SlideToGenerate { PivotType = 2, SlideTemplateIndex = SLIDE_TEMPLATE_2_INDEX },
            //    new SlideToGenerate { PivotType = 1, SlideTemplateIndex = SLIDE_TEMPLATE_2_INDEX },

            //    new SlideToGenerate { PivotType = 1, SlideTemplateIndex = SLIDE_TEMPLATE_1_INDEX },
            //    new SlideToGenerate { PivotType = 2, SlideTemplateIndex = SLIDE_TEMPLATE_2_INDEX },
            //    new SlideToGenerate { PivotType = 3, SlideTemplateIndex = SLIDE_TEMPLATE_3_INDEX },
            //    new SlideToGenerate { PivotType = 4, SlideTemplateIndex = SLIDE_TEMPLATE_4_INDEX },


            //    //new SlideToGenerate { PivotType = 3, SlideIndex = 4 },
            //    //new SlideToGenerate { PivotType = 3, SlideIndex = 5 },
            //    //new SlideToGenerate { PivotType = 3, SlideIndex = 6 },
            //    //new SlideToGenerate { PivotType = 4, SlideIndex = 7 },
            //    //new SlideToGenerate { PivotType = 4, SlideIndex = 8 },
            //    //new SlideToGenerate { PivotType = 4, SlideIndex = 9 },
            //    //new SlideToGenerate { PivotType = 4, SlideIndex = 10 }
            //};
        }

        static void PredisposiTmpFolder()
        {
            // Todo: usare qualcosa tipo AppDomain.CurrentDomain.BaseDirectory;
            // context.TmpFolder = AppDomain.CurrentDomain.BaseDirectory + "\\tmp";
            context.TmpFolder = context.OutputFolder + "\\tmp";

            // Pulizia e creazione della cartella temporanea
            if (Directory.Exists(context.TmpFolder)) { Directory.Delete(context.TmpFolder, true); }
            Directory.CreateDirectory(context.TmpFolder);
        }

        static void GeneraFileConDati()
        {
            // Simulazione del creazione del file Excel con i dati...
            context.ExcelDataSourceFile = context.TmpFolder + "\\DataSource.xlsx";
            File.Copy(XLS_DATASOURCE_FILE_PATH, context.ExcelDataSourceFile, true);
        }


        static void CreaListaImmaginiDaGenerare()
        {
            string percorsoFile = PPT_PRINT_AREAS_FILE_PATH;

            if (File.Exists(percorsoFile))
            {
                context.ItemsToExportAsImage = new List<ItemToExport>();

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

                    context.ItemsToExportAsImage.Add(new ItemToExport { ImageId = imageId, Sheet = sheet, PrintArea = printArea, });
                }
            }
            else
            {
                //todo: eccezione managed
                Console.WriteLine("Il file non esiste!");
            }
        }

        static void GeneraimmaginiPivotTables()
        {
            // Carica il workbook
            Workbook workbook = new Workbook(context.ExcelDataSourceFile);

            const int PIVOT_TYPES_NUMBER = 4;

            // context.GeneratedImagesList = new List<GeneratedImages>();

            for (int pivotType = 1; pivotType <= PIVOT_TYPES_NUMBER; pivotType++)
            {
                var worksheetName = $"Shapes_{pivotType.ToString("D2")}";
                var worksheetWithPivot = workbook.Worksheets[worksheetName];

                // Prendi la prima PivotTable del foglio
                PivotTable pivot = worksheetWithPivot.PivotTables[0];

                // Ottieni il range della pivot
                CellArea range = pivot.TableRange2;

                // Imposta l'area di stampa sul range della pivot
                worksheetWithPivot.PageSetup.PrintArea =
                    CellsHelper.CellIndexToName(range.StartRow, range.StartColumn) + ":" +
                    CellsHelper.CellIndexToName(range.EndRow, range.EndColumn);

                // Opzioni di rendering in immagine
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    OnePagePerSheet = true,
                    PrintingPage = PrintingPageType.Default
                };
                // Crea un oggetto SheetRender per il foglio
                SheetRender sr = new SheetRender(worksheetWithPivot, imgOptions);
                //   sr = new SheetRender(worksheetWithPivot, imgOptions);

                // Esporta la pivot (tutto il foglio) come immagine
                var imagePath = context.TmpFolder + $"\\Img_{worksheetName}_Pivot.png";
                sr.ToImage(0, imagePath);

                // aggiungo alla lista delle immagini generate
                //context.GeneratedImagesList.Add(new GeneratedImages { Path = imagePath, PivotType = pivotType });
            }
        }


        static void GeneraimmaginiItmes()
        {
            // Carica il workbook
            Workbook workbook = new Workbook(context.ExcelDataSourceFile);

            foreach (var itemsToExportAsImage in context.ItemsToExportAsImage)
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
                ChopImage(toBeChoppedImagePath, finalImagePath);

                //todo uncommentare
                //File.Delete(toBeChoppedImagePath);
            }
        }

        static private string GetImagePath(string imageId)
        {
            var imagePath = $"{context.TmpFolder}\\{imageId}.png";
            return imagePath;
        }

        static void ChopImage(string inputPath, string outputPath)
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


        static void CreazionePowerPoint()
        {
            context.PowerPointOutputFile = context.OutputFolder + "\\" + OUTPUT_POWERPOINT_FILENAME;



            //   const string OUTPUT_FILE = @"C:\Data\Lefo\Dev\GDPptxReport\POC\Output.pptx";

            // ripulisco il possibile file di output
            if (File.Exists(context.PowerPointOutputFile))
            { File.Delete(context.PowerPointOutputFile); }


            // apro la presentazione
            //var pres = new Presentation(TEMPLATE_FILE);
            //pres.Save(OUTPUT_FILE);
            //pres = new Presentation(OUTPUT_FILE);

            // copio il template nella cartella di output
            File.Copy(PPT_TEMPLATE_FILE_PATH, context.PowerPointOutputFile);
            var pres = new ShapeCrawler.Presentation(context.PowerPointOutputFile);

            // prendo la slide con 2 contenuti verticali
            //var slideToDuplicate = pres.Slide(SLIDE_INDEX_CONTENT_2);



            const int SpazionIntornoAlleImmagini = 10;
            // const int DistanzaDallaBordoLaterale = 50;

            foreach (var slideToGenerate in context.SlideToGenerateList)
            {
                //TODO: inserire qui il codice per generare i template in 
                // creo le slides a partire dai template
                pres.Slides.Add(pres.Slide(slideToGenerate.SlideType), pres.Slides.Count + 1);
                var slideToEdit = pres.Slide(pres.Slides.Count);




                // in base al template inserisco una o più immagini in diverse posizioni
                var imgFilePath = GetImagePath(slideToGenerate.ImageId1);
                var imgStream = new FileStream(imgFilePath, FileMode.Open, FileAccess.Read);
                decimal larghezzaImmagine;
                decimal altezzaImmagine;
                int numeroImmaginiInOrizzontale;
                int numeroImmaginiInVerticale;
                switch (slideToGenerate.SlideType)
                {
                    case 1:
                        // un'unica immagine che occupa tutta la slide
                        numeroImmaginiInVerticale = 1;
                        numeroImmaginiInOrizzontale = 1;
                        larghezzaImmagine = pres.SlideWidth / numeroImmaginiInOrizzontale - SpazionIntornoAlleImmagini * 2;
                        altezzaImmagine = pres.SlideHeight / numeroImmaginiInVerticale - SpazionIntornoAlleImmagini * 2;
                        //
                        slideToEdit.Shapes.AddPicture(imgStream);
                        slideToEdit.Shapes[1].Y = SpazionIntornoAlleImmagini;
                        slideToEdit.Shapes[1].X = SpazionIntornoAlleImmagini;
                        slideToEdit.Shapes[1].Width = larghezzaImmagine;
                        slideToEdit.Shapes[1].Height = altezzaImmagine;
                        break;
                    case 2:
                        // 2 immagini sulla stessa riga
                        numeroImmaginiInVerticale = 1;
                        numeroImmaginiInOrizzontale = 2;
                        larghezzaImmagine = pres.SlideWidth / numeroImmaginiInOrizzontale - SpazionIntornoAlleImmagini * 2;
                        altezzaImmagine = pres.SlideHeight / numeroImmaginiInVerticale - SpazionIntornoAlleImmagini * 2;
                        //
                        slideToEdit.Shapes.AddPicture(imgStream);
                        slideToEdit.Shapes[1].Y = SpazionIntornoAlleImmagini;
                        slideToEdit.Shapes[1].X = SpazionIntornoAlleImmagini;
                        slideToEdit.Shapes[1].Width = larghezzaImmagine;
                        slideToEdit.Shapes[1].Height = altezzaImmagine;
                        //
                        imgFilePath = GetImagePath(slideToGenerate.ImageId2);
                        var imgStream2 = new FileStream(imgFilePath, FileMode.Open, FileAccess.Read);
                        slideToEdit.Shapes.AddPicture(imgStream);
                        slideToEdit.Shapes[2].Y = SpazionIntornoAlleImmagini;
                        slideToEdit.Shapes[2].X = larghezzaImmagine + (SpazionIntornoAlleImmagini * 3);
                        slideToEdit.Shapes[2].Width = larghezzaImmagine;
                        slideToEdit.Shapes[2].Height = altezzaImmagine;
                        break;
                    case 3:
                        //var imgFilePath = GetImagePath(slideToGenerate.ImageId1);
                        //var imgStream = new FileStream(imgFilePath, FileMode.Open, FileAccess.Read);
                        //slideToEdit.Shapes.AddPicture(imgStream);
                        //slideToEdit.Shapes[1].Y = 50;
                        //slideToEdit.Shapes[1].X = 50;
                        //slideToEdit.Shapes[1].Width = pres.SlideWidth - 100;
                        //slideToEdit.Shapes[1].Height = pres.SlideHeight - 100;
                        break;
                    case 4:
                        //var imgFilePath = GetImagePath(slideToGenerate.ImageId1);
                        //var imgStream = new FileStream(imgFilePath, FileMode.Open, FileAccess.Read);
                        //slideToEdit.Shapes.AddPicture(imgStream);
                        //slideToEdit.Shapes[1].Y = 50;
                        //slideToEdit.Shapes[1].X = 50;
                        //slideToEdit.Shapes[1].Width = pres.SlideWidth - 100;
                        //slideToEdit.Shapes[1].Height = pres.SlideHeight - 100;
                        break;
                    default:
                        break;
                }
            }

            // rimuovo i 4 templale
            for (var slideIndex = 1; slideIndex <= NumberOfTemplateSlides; slideIndex++)
            { pres.Slide(1).Remove(); }

            pres.Save(context.PowerPointOutputFile);
        }

        //static void AggiungiImmagini()
        //{
        //    context.PowerPointOutputFile = context.TmpFolder + "\\Output_AggiuntaImmagini.pptx";

        //    var pres = new Presentation(context.PowerPointOutputFile);
        //    var slideToCheck = pres.Slide(1);

        //    var imgStream = new FileStream(@"C:\Data\Lefo\Dev\GDPptxReport\POC\Images\foto_01.png", FileMode.Open, FileAccess.Read);
        //    slideToCheck.Shapes.AddPicture(imgStream);

        //    pres.Save(context.PowerPointOutputFile);
        //}
    }
}