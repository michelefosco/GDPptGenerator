using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Pivot;
using Aspose.Cells.Rendering;
using DocumentFormat.OpenXml.Office.CoverPageProps;
using ImageMagick;
using ShapeCrawler;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace POC_Console
{
    internal class Program
    {

        const int NumberOfTemplateSlides = 4;

        //todo: da mofificare con i percorsi reali, prendere da cartella di esecuzione

        const string TEMPLATE_FILE_PPT = @"C:\Data\Lefo\Dev\GDPptGenerator\Solution\Templates\Template_PowerPoint.pptx";
        const string TEMPLATE_FILE_XLS = @"C:\Data\Lefo\Dev\GDPptGenerator\Solution\Templates\Template_DataSource.xlsx";

        const string OUTPUT_POWERPOINT_FILENAME = "Output.pptx";

        //static string foder = AppDomain.CurrentDomain.BaseDirectory;


        struct Context
        {
            public string OutputFolder;
            public string TmpFolder;
            public string ExcelDataSourceFile;
            public string PowerPointOutputFile;
            public List<SlideToGenerate> SlideToGenerateList;
            public List<GeneratedImages> GeneratedImagesList;
        }

        struct GeneratedImages
        {
            public int PivotType;
            public string Path;
        }

        struct SlideToGenerate
        {
            public int PivotType;
            public int SlideTemplateIndex;
        }

        static Context context;

        static void Main()
        {
            // scelta dall'utente...
            context.OutputFolder = @"C:\Data\Lefo\Dev\GDPptGenerator\AAA_Output";

            CreaListaSlidesDaGenerare();

            PredisposiTmpFolder();

            GeneraFileConDati();

            GeneraimmaginiPivotTables();

            CreazionePowerPoint();

            //      AggiungiImmagini();
        }

        static void CreaListaSlidesDaGenerare()
        {
            const int SLIDE_TEMPLATE_1_INDEX = 1;
            const int SLIDE_TEMPLATE_2_INDEX = 2;
            const int SLIDE_TEMPLATE_3_INDEX = 3;
            const int SLIDE_TEMPLATE_4_INDEX = 4;
            context.SlideToGenerateList = new List<SlideToGenerate>
            {
                new SlideToGenerate { PivotType = 4, SlideTemplateIndex = SLIDE_TEMPLATE_1_INDEX },
                new SlideToGenerate { PivotType = 3, SlideTemplateIndex = SLIDE_TEMPLATE_1_INDEX },
                new SlideToGenerate { PivotType = 2, SlideTemplateIndex = SLIDE_TEMPLATE_1_INDEX },
                new SlideToGenerate { PivotType = 1, SlideTemplateIndex = SLIDE_TEMPLATE_1_INDEX },

                new SlideToGenerate { PivotType = 4, SlideTemplateIndex = SLIDE_TEMPLATE_2_INDEX },
                new SlideToGenerate { PivotType = 3, SlideTemplateIndex = SLIDE_TEMPLATE_2_INDEX },
                new SlideToGenerate { PivotType = 2, SlideTemplateIndex = SLIDE_TEMPLATE_2_INDEX },
                new SlideToGenerate { PivotType = 1, SlideTemplateIndex = SLIDE_TEMPLATE_2_INDEX },

                new SlideToGenerate { PivotType = 1, SlideTemplateIndex = SLIDE_TEMPLATE_1_INDEX },
                new SlideToGenerate { PivotType = 2, SlideTemplateIndex = SLIDE_TEMPLATE_2_INDEX },
                new SlideToGenerate { PivotType = 3, SlideTemplateIndex = SLIDE_TEMPLATE_3_INDEX },
                new SlideToGenerate { PivotType = 4, SlideTemplateIndex = SLIDE_TEMPLATE_4_INDEX },


                //new SlideToGenerate { PivotType = 3, SlideIndex = 4 },
                //new SlideToGenerate { PivotType = 3, SlideIndex = 5 },
                //new SlideToGenerate { PivotType = 3, SlideIndex = 6 },
                //new SlideToGenerate { PivotType = 4, SlideIndex = 7 },
                //new SlideToGenerate { PivotType = 4, SlideIndex = 8 },
                //new SlideToGenerate { PivotType = 4, SlideIndex = 9 },
                //new SlideToGenerate { PivotType = 4, SlideIndex = 10 }
            };
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
            File.Copy(TEMPLATE_FILE_XLS, context.ExcelDataSourceFile, true);
        }


        static void GeneraimmaginiPivotTables()
        {
            // Carica il workbook
            Workbook workbook = new Workbook(context.ExcelDataSourceFile);

            const int PIVOT_TYPES_NUMBER = 4;

            context.GeneratedImagesList = new List<GeneratedImages>();

            for (int pivotType = 1; pivotType <= PIVOT_TYPES_NUMBER; pivotType++)
            {
                var worksheetName = $"Pivot_{pivotType.ToString("D2")}";
                var worksheetWithPivot = workbook.Worksheets[worksheetName];

                // Trova tutti i fogli che iniziano con "Pivot_"
                //   var worksheetsWithPivots = workbook.Worksheets.Where(_ => _.Name.StartsWith("Pivot_", StringComparison.InvariantCultureIgnoreCase)).ToList();
                //foreach (var worksheetWithPivot in worksheetsWithPivots)
                //{
                
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
                var imagePath = context.TmpFolder + $"\\Img_{worksheetName}.png";
                sr.ToImage(0, imagePath);

                // aggiungo alla lista delle immagini generate
                context.GeneratedImagesList.Add(new GeneratedImages { Path = imagePath, PivotType = pivotType });
            }
        }




        static void CreazionePowerPoint()
        {
            context.PowerPointOutputFile = context.OutputFolder + "\\" + OUTPUT_POWERPOINT_FILENAME;



            //   const string OUTPUT_FILE = @"C:\Data\Lefo\Dev\GDPptxReport\POC\Output.pptx";

            // ripulisco il possibile file di output
            if (File.Exists(context.PowerPointOutputFile))
            { File.Delete(context.PowerPointOutputFile); }
            File.Copy(TEMPLATE_FILE_PPT, context.PowerPointOutputFile);

            // apro la presentazione
            //var pres = new Presentation(TEMPLATE_FILE);
            //pres.Save(OUTPUT_FILE);
            //pres = new Presentation(OUTPUT_FILE);

            var pres = new Presentation(context.PowerPointOutputFile);

            // prendo la slide con 2 contenuti verticali
            //var slideToDuplicate = pres.Slide(SLIDE_INDEX_CONTENT_2);

            foreach (var slideToGenerate in context.SlideToGenerateList)
            {
                //TODO: inserire qui il codice per generare i template in 
                // creo le slides a partire dai template
                pres.Slides.Add(pres.Slide(slideToGenerate.SlideTemplateIndex), pres.Slides.Count + 1);

                var slideToEdit = pres.Slide(pres.Slides.Count);
                var imgFilePath = context.GeneratedImagesList.Single(_ => _.PivotType == slideToGenerate.PivotType).Path;
                var imgStream = new FileStream(imgFilePath, FileMode.Open, FileAccess.Read);
                slideToEdit.Shapes.AddPicture(imgStream);
                slideToEdit.Shapes[1].Y = 50;
                slideToEdit.Shapes[1].X = 50;
                slideToEdit.Shapes[1].Width = pres.SlideWidth - 100;
                slideToEdit.Shapes[1].Height = pres.SlideHeight - 100;

                //pres.Slides.Add(pres.Slide(SLIDE_TEMPLATE_2_INDEX), pres.Slides.Count + 1);
                //pres.Slides.Add(pres.Slide(SLIDE_TEMPLATE_2_INDEX), pres.Slides.Count + 1);
                //pres.Slides.Add(pres.Slide(SLIDE_TEMPLATE_3_INDEX), pres.Slides.Count + 1);
                //pres.Slides.Add(pres.Slide(SLIDE_TEMPLATE_3_INDEX), pres.Slides.Count + 1);
                //pres.Slides.Add(pres.Slide(SLIDE_TEMPLATE_3_INDEX), pres.Slides.Count + 1);
                //pres.Slides.Add(pres.Slide(SLIDE_TEMPLATE_4_INDEX), pres.Slides.Count + 1);
                //pres.Slides.Add(pres.Slide(SLIDE_TEMPLATE_4_INDEX), pres.Slides.Count + 1);
                //pres.Slides.Add(pres.Slide(SLIDE_TEMPLATE_4_INDEX), pres.Slides.Count + 1);
                //pres.Slides.Add(pres.Slide(SLIDE_TEMPLATE_4_INDEX), pres.Slides.Count + 1);
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