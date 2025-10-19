using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using FilesEditor.Entities;
using FilesEditor.Enums;
using ShapeCrawler;
using System;
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
            creaFiles_Presentazioni();
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }


        private void creaFiles_Presentazioni()
        {
            var outputFileNames = Context.SildeToGenerate.Select(s => s.OutputFileName).Distinct().ToList();

            foreach (var outputFileName in outputFileNames)
            {
                var slideToGenerateList = Context.SildeToGenerate.Where(s => s.OutputFileName == outputFileName).ToList();

                #region Ottengo i percorsi dei file
                var outputfilePath = GetOutputFilePath(Context.DestinationFolder, outputFileName);
                var percorsoFileTemplatePowerPoint = GetTemplateFilePath(Context.DataSourceFilePath);
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
                int SLIDE_TEMPLATE_POSITION = pres.Slides.Count; // la slide template è l'ultima del file
                for (int j = 1; j <= slideToGenerateList.Count - 1; j++)
                {
                    pres.Slides.Add(pres.Slide(SLIDE_TEMPLATE_POSITION), pres.Slides.Count + 1);
                }
                #endregion


                var pizexlPerCm = 28.35;

                const int spazioneDiSeparazioneOrizzontaleTraDueImmagini = 80;
                const int spazioneDiSeparazioneVerticaleTraDueImmagini = 50;
                //
                int offSetVerticale = (int)(2.59 * pizexlPerCm);
                int offSetOrizzontale = (int)(1.05 * pizexlPerCm);
                //
                var totalAvailableWidth = (int)(27.4 * pizexlPerCm);
                var totalAvailableHeight = (int)(12.6 * pizexlPerCm);

                int slideToEditIndex = SLIDE_TEMPLATE_POSITION;
                foreach (var slideToGenerate in slideToGenerateList)
                {
                    var slideToEdit = pres.Slide(slideToEditIndex);

                    #region Modifico la textbox del titolo
                    var titleTextBox = slideToEdit.GetTextBoxes().FirstOrDefault(tb => tb.Text.Contains("Titolo"));
                    if (titleTextBox != null)
                    {
                        titleTextBox.SetText(slideToGenerate.Title);
                    }
                    #endregion

                    #region Aggiungo le immagini in base al tipo di slide
                    // dimensioni del riquadro contenente un immagine
                    decimal boxWidth;
                    decimal boxHeight;
                    //
                    decimal boxPostionX;
                    decimal boxPostionY;

                    // somma degli spazi usati per separare le immagini
                    int spaceUsedToSeparateImages;

                    var numeroImmagini = slideToGenerate.Contents.Count();

                    switch (numeroImmagini)
                    {
                        case 1:
                            // Spazio massimo a disposizione per le immagini
                            boxWidth = totalAvailableWidth;
                            boxHeight = totalAvailableHeight;
                            // posizione del box
                            boxPostionY = offSetVerticale;
                            boxPostionX = offSetOrizzontale;

                            AddImageToTheSlide(slide: slideToEdit,
                                            imageId: slideToGenerate.Contents[0],
                                        boxWidth: boxWidth,
                                        boxHeight: boxHeight,
                                        boxPostionX: boxPostionX,
                                        boxPostionY: boxPostionY
                                        );
                            break;

                        case 2:
                        case 3:
                            if (slideToGenerate.LayoutType == LayoutTypes.Horizontal)
                            {
                                // spazio totale usato per lo spazio tra le immagini
                                spaceUsedToSeparateImages = spazioneDiSeparazioneOrizzontaleTraDueImmagini * (numeroImmagini - 1);

                                // Spazio massimo a disposizione per le immagini
                                boxWidth = (totalAvailableWidth - spaceUsedToSeparateImages) / numeroImmagini;
                                boxHeight = totalAvailableHeight;

                                for (var imagePosition = 0; imagePosition < numeroImmagini; imagePosition++)
                                {
                                    // posizione del box
                                    boxPostionX = offSetOrizzontale + (boxWidth * imagePosition) + (spazioneDiSeparazioneOrizzontaleTraDueImmagini * imagePosition);
                                    boxPostionY = offSetVerticale; // costante per tutte le immagini

                                    AddImageToTheSlide(
                                        slide: slideToEdit,
                                        imageId: slideToGenerate.Contents[imagePosition],
                                        boxWidth: boxWidth,
                                        boxHeight: boxHeight,
                                        boxPostionX: boxPostionX,
                                        boxPostionY: boxPostionY
                                        );
                                }
                            }
                            else
                            {
                                // spazio totale usato per lo spazio tra le immagini
                                spaceUsedToSeparateImages = spazioneDiSeparazioneVerticaleTraDueImmagini * (numeroImmagini - 1);

                                // Spazio massimo a disposizione per le immagini
                                boxWidth = totalAvailableWidth;
                                boxHeight = (totalAvailableHeight - spaceUsedToSeparateImages) / numeroImmagini;

                                for (var imagePosition = 0; imagePosition < numeroImmagini; imagePosition++)
                                {
                                    // posizione del box
                                    boxPostionX = offSetOrizzontale; // costante per tutte le immagini
                                    boxPostionY = offSetVerticale + (boxHeight * imagePosition) + (spazioneDiSeparazioneVerticaleTraDueImmagini * imagePosition);

                                    AddImageToTheSlide(
                                        slide: slideToEdit,
                                        imageId: slideToGenerate.Contents[imagePosition],
                                        boxWidth: boxWidth,
                                        boxHeight: boxHeight,
                                        boxPostionX: boxPostionX,
                                        boxPostionY: boxPostionY
                                        );
                                }
                            }
                            break;

                        default:
                            throw new ArgumentOutOfRangeException("Numero di immagini per slide non gestito");
                    }
                    #endregion

                    slideToEditIndex++;
                }

                #region Operazioni finali
                // rimuovo le slide template
                //for (var slideIndex = 1; slideIndex <= NumberOfTemplateSlides; slideIndex++)
                //{ pres.Slide(SLIDE_TEMPLATE_1_INDEX).Remove(); }

                // rimuovo le slide vuote finali
                //int numberOfSlidesToRemove = pres.Slides.Count - (slideToEditIndex - 1);
                //for (var j = 1; j <= numberOfSlidesToRemove; j++)
                //{
                //    // rimuovo l'ultima slide
                //    pres.Slide(pres.Slides.Count).Remove();
                //}

                pres.Save(outputfilePath);

                Context.OutputFilePathLists.Add(outputfilePath);
                #endregion
            }
        }

        private IShape AddImageToTheSlide(ISlide slide, string imageId, decimal boxWidth, decimal boxHeight, decimal boxPostionY, decimal boxPostionX)
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


        private string GetTemplateFilePath(string dataSourceFilePath)
        {
            var sourceFilesFolder = System.IO.Path.GetDirectoryName(Context.DataSourceFilePath);
            var percorsoFileTemplatePowerPoint = System.IO.Path.Combine(sourceFilesFolder, Constants.FileNames.POWERPOINT_TEMPLATE_FILENAME);
            return percorsoFileTemplatePowerPoint;
        }
        private string GetOutputFilePath(string destinationFolder, string outputFileName)
        {
            var outputfilePath = System.IO.Path.Combine(Context.DestinationFolder, outputFileName);
            if (outputfilePath.EndsWith(".pptx", StringComparison.InvariantCultureIgnoreCase) == false)
            { outputfilePath = outputFileName + ".pptx"; }
            return outputfilePath;
        }
    }
}