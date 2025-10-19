using DocumentFormat.OpenXml.Office2010.ExcelAc;
using FilesEditor.Entities;
using FilesEditor.Enums;
using ShapeCrawler;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace FilesEditor.Steps.BuildPresentation
{
    internal class Step_CreaFiles_Presentazioni : StepBase
    {
        public Step_CreaFiles_Presentazioni(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            Context.DebugInfoLogger.LogStepContext("Step_CreaFiles_Presentazioni", Context);
            buildPresentations();
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }


        private void buildPresentations()
        {
            var outputFileNames = Context.SildeToGenerate.Select(s => s.OutputFileName).Distinct().ToList();

            foreach (var outputFileName in outputFileNames)
            {
                var slideToGenerateList = Context.SildeToGenerate.Where(s => s.OutputFileName == outputFileName).ToList();
                buildPresentation(outputFileName, slideToGenerateList);
            }
        }

        private void buildPresentation(string outputFileName, List<SlideToGenerate> slideToGenerateList)
        {
            //var slideToGenerateList = Context.SildeToGenerate.Where(s => s.OutputFileName == outputFileName).ToList();

            #region Ottengo i percorsi dei file
            var outputfilePath = getOutputFilePath(Context.DestinationFolder, outputFileName);
            var percorsoFileTemplatePowerPoint = getTemplateFilePath(Context.DataSourceFilePath);
            #endregion

            #region Copio il template nella cartella di output
            // ripulisco il possibile file di output
            CancellaFileSeEsiste(outputfilePath, FileTypes.PresentationOutput);
            File.Copy(percorsoFileTemplatePowerPoint, outputfilePath);
            #endregion

            var pres = new ShapeCrawler.Presentation(outputfilePath);

            #region Preparo la slide "indice" con la lista dei titoli delle slides
            int SLIDE_INDEX_POSITION = 2;
            var slide = pres.Slide(SLIDE_INDEX_POSITION);
            var titlesListBox = slide.GetTextBoxes().LastOrDefault();
            var titleIndex = 1;
            foreach (var slideToGenerate in slideToGenerateList)
            {
                titlesListBox.Paragraphs.Add(slideToGenerate.Title, titleIndex);
                titleIndex++;
            }
            titlesListBox.Paragraphs[0].Remove(); // rimuovo il testo di default
            #endregion

            #region predispongo le slides duplicando quella template (ovvero l'ultima del file)
            var SLIDE_TEMPLATE_POSITION = pres.Slides.Count; // la slide template è l'ultima del file
            for (var j = 1; j <= slideToGenerateList.Count - 1; j++)
            { pres.Slides.Add(pres.Slide(SLIDE_TEMPLATE_POSITION), pres.Slides.Count + 1); }
            #endregion

            // rapporto tra numero di pixel e cm
            var pizexlPerCm = 28.35;
            //
            int offSetVerticale = (int)(2.59 * pizexlPerCm);
            int offSetOrizzontale = (int)(1.05 * pizexlPerCm);
            //
            var totalAvailableWidth = (int)(27.4 * pizexlPerCm);
            var totalAvailableHeight = (int)(12.6 * pizexlPerCm);
            //
            const int spazioOrizzontaleTraDueImmagini = 2;
            const int spazioVerticaleTraDueImmagini = 1;

            int slideToEditIndex = SLIDE_TEMPLATE_POSITION;
            foreach (var slideToGenerate in slideToGenerateList)
            {
                var slideToEdit = pres.Slide(slideToEditIndex);

                #region Modifico la textbox del titolo
                var titleTextBox = slideToEdit.GetTextBoxes().FirstOrDefault(tb => tb.Text.Contains("Titolo"));
                if (titleTextBox != null)
                { titleTextBox.SetText(slideToGenerate.Title); }
                #endregion

                // numero di immagini da collocare
                var numOfImages = slideToGenerate.Contents.Count();

                #region dimensione dei box
                decimal boxWidth = totalAvailableWidth;
                decimal boxHeight = totalAvailableHeight;
                if (numOfImages > 1)
                {
                    if (slideToGenerate.LayoutType == LayoutTypes.Horizontal)
                    {
                        // la larghezza del box è data dal totale disponibile meno tutto quello usato per lo spazio orizzontale lasciato tra le immagini il tutto diviso per numero di immagini
                        boxWidth = (totalAvailableWidth - (spazioOrizzontaleTraDueImmagini * (numOfImages - 1))) / numOfImages;
                    }
                    else
                    {
                        // l'altezza del box è data dal totale disponibile meno tutto quello usato per lo spazio verticale tra lasciato le immagini il tutto diviso per numero di immagini
                        boxHeight = (totalAvailableHeight - (spazioVerticaleTraDueImmagini * (numOfImages - 1))) / numOfImages;
                    }
                }
                #endregion


                for (var imagePosition = 0; imagePosition < numOfImages; imagePosition++)
                {
                    #region posizione del box
                    decimal boxPostionY = offSetVerticale;
                    decimal boxPostionX = offSetOrizzontale;
                    if (numOfImages > 1)
                    {
                        if (slideToGenerate.LayoutType == LayoutTypes.Horizontal)
                        { boxPostionX = offSetOrizzontale + (boxWidth * imagePosition) + (spazioOrizzontaleTraDueImmagini * imagePosition); }
                        else
                        { boxPostionY = offSetVerticale + (boxHeight * imagePosition) + (spazioVerticaleTraDueImmagini * imagePosition); }
                    }
                    #endregion

                    addImageToTheSlide(
                        slide: slideToEdit,
                        imageId: slideToGenerate.Contents[imagePosition],
                        boxWidth: boxWidth,
                        boxHeight: boxHeight,
                        boxPostionX: boxPostionX,
                        boxPostionY: boxPostionY
                        );
                }

                slideToEditIndex++;
            }

            #region Operazioni finali
            pres.Save(outputfilePath);
            Context.OutputFilePathLists.Add(outputfilePath);
            #endregion
        }



        private IShape addImageToTheSlide(ISlide slide, string imageId, decimal boxWidth, decimal boxHeight, decimal boxPostionY, decimal boxPostionX)
        {
            var imgFilePath = GetTmpFolderImagePathByImageId(Context.TmpFolder, imageId);
            var imgStream = new FileStream(imgFilePath, FileMode.Open, FileAccess.Read);
            slide.Shapes.AddPicture(imgStream);
            imgStream.Close();

            using (Image img = Image.FromFile(imgFilePath))
            {
                Console.WriteLine($"Width: {img.Width}px");
                Console.WriteLine($"Height: {img.Height}px");
            }

            var shape = slide.Shapes[slide.Shapes.Count - 1];
            shape.Width = boxWidth;
            shape.Height = boxHeight;
            shape.Y = boxPostionY;
            shape.X = boxPostionX;

            return shape;
        }
        private string getTemplateFilePath(string dataSourceFilePath)
        {
            var sourceFilesFolder = System.IO.Path.GetDirectoryName(Context.DataSourceFilePath);
            var percorsoFileTemplatePowerPoint = System.IO.Path.Combine(sourceFilesFolder, Constants.FileNames.POWERPOINT_TEMPLATE_FILENAME);
            return percorsoFileTemplatePowerPoint;
        }
        private string getOutputFilePath(string destinationFolder, string outputFileName)
        {
            var outputfilePath = System.IO.Path.Combine(Context.DestinationFolder, outputFileName);
            if (outputfilePath.EndsWith(".pptx", StringComparison.InvariantCultureIgnoreCase) == false)
            { outputfilePath = outputFileName + ".pptx"; }
            return outputfilePath;
        }
    }
}