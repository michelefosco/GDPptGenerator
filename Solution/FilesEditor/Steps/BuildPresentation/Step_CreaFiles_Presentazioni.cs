using DocumentFormat.OpenXml.Vml;
using FilesEditor.Entities;
using FilesEditor.Enums;
using ShapeCrawler;
using System;
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

                #region Predispongo il file di output
                var outputfilePath = System.IO.Path.Combine(Context.DestinationFolder, outputFileName);
                if (outputfilePath.EndsWith(".pptx", StringComparison.InvariantCultureIgnoreCase) == false)
                { outputfilePath = outputFileName + ".pptx"; }

                // ripulisco il possibile file di output
                CancellaFileSeEsiste(outputfilePath, FileTypes.PresentationOutput);
                #endregion


                #region Copio il template nella cartella di output
                var sourceFilesFolder = System.IO.Path.GetDirectoryName(Context.DataSourceFilePath);
                var percorsoFileTemplatePowerPoint = System.IO.Path.Combine(sourceFilesFolder, Constants.FileNames.POWERPOINT_TEMPLATE_FILENAME);
                File.Copy(percorsoFileTemplatePowerPoint, outputfilePath);
                var pres = new ShapeCrawler.Presentation(outputfilePath);
                #endregion

                //todo: creazione slide "Indice"
                int SLIDE_INDEX_POSITION = 2;
                var titleIndex = 1;
                var slide = pres.Slide(SLIDE_INDEX_POSITION);
                var titlesListBox = slide.GetTextBoxes().LastOrDefault();
                foreach (var slideToGenerate in slideToGenerateList)
                {
                    titlesListBox.Paragraphs.Add(slideToGenerate.Title, titleIndex);
                    titleIndex++;

                    //slide.Shapes.AddShape(100, 100, 400, 200,Geometry.Rectangle);
                    //var autoShape = slide.Shapes[slide.Shapes.Count - 1];// as IAutoShape;
                    //var paragraphCollection = autoShape.TextFrame.Paragraphs;

                    //// Add first bullet point
                    //var p1 = paragraphCollection.Add();
                    //p1.Text = "First bullet item";
                    //p1.TextFormat.Bullet.Type = BulletType.None; // Enable bullet
                    //p1.TextFormat.Bullet.Character = '•'; // Bullet symbol

                    //// Add second bullet point
                    //var p2 = paragraphCollection.Add();
                    //p2.Text = "Second bullet item";
                    //p2.TextFormat.Bullet.Type = BulletType.None;
                    //p2.TextFormat.Bullet.Character = '•';

                    //// Add third bullet point
                    //var p3 = paragraphCollection.Add();
                    //p3.Text = "Third bullet item";
                    //p3.TextFormat.Bullet.Type = BulletType.None;
                    //p3.TextFormat.Bullet.Character = '•';
                }
                titlesListBox.Paragraphs[0].Remove(); // rimuovo il testo di default

                #region predispongo le slides duplicando quella template (ovvero l'ultima del file)
                int SLIDE_TEMPLATE_POSITION = pres.Slides.Count; // la slide template è l'ultima del file
                for (int j = 1; j <= slideToGenerateList.Count - 1; j++)
                {
                    pres.Slides.Add(pres.Slide(SLIDE_TEMPLATE_POSITION), pres.Slides.Count + 1);
                }
                #endregion

                const int SpazionIntornoAlleImmagini = 10;
                const int offSetVerticale = 80;
                int slideToEditIndex = SLIDE_TEMPLATE_POSITION;
                foreach (var slideToGenerate in slideToGenerateList)
                {
                    //#region Duplico la slide template
                    //pres.Slides.Add(pres.Slide(SLIDE_TEMPLATE_1_INDEX), pres.Slides.Count + 1);
                    //var slideToEdit = pres.Slide(pres.Slides.Count);
                    //#endregion

                    var slideToEdit = pres.Slide(slideToEditIndex);
                    #region Modifico la textbox del titolo
                    var titleTextBox = slideToEdit.GetTextBoxes().FirstOrDefault(tb => tb.Text.Contains("Titolo"));
                    if (titleTextBox != null)
                    {
                        titleTextBox.SetText(slideToGenerate.Title);
                    }
                    #endregion


                    #region Aggiungo le immagini in base al tipo di slide
                    decimal imageWidth;
                    decimal imageHeight;
                    int numeroImmaginiInOrizzontale;
                    int numeroImmaginiInVerticale;

                    if (slideToGenerate.LayoutType == LayoutTypes.Horizontal)
                    {
                        switch (slideToGenerate.Contents.Count())
                        {
                            case 1:
                                // un'unica immagine che occupa tutta la slide
                                numeroImmaginiInVerticale = 1;
                                numeroImmaginiInOrizzontale = 1;
                                imageWidth = pres.SlideWidth / numeroImmaginiInOrizzontale - SpazionIntornoAlleImmagini * 2;
                                imageHeight = (pres.SlideHeight - offSetVerticale) / numeroImmaginiInVerticale - SpazionIntornoAlleImmagini * 2;
                                AddImageToTheSlide(slide: slideToEdit,
                                                imageId: slideToGenerate.Contents[0],
                                                imageWidth: imageWidth,
                                                imageHeight: imageHeight,
                                                imagePostionY: offSetVerticale + SpazionIntornoAlleImmagini,
                                                imagePostionX: SpazionIntornoAlleImmagini);
                                break;

                            case 2:
                                // 2 immagini sulla stessa riga
                                numeroImmaginiInVerticale = 1;
                                numeroImmaginiInOrizzontale = 2;
                                imageWidth = pres.SlideWidth / numeroImmaginiInOrizzontale - SpazionIntornoAlleImmagini * 2;
                                imageHeight = (pres.SlideHeight - offSetVerticale) / numeroImmaginiInVerticale - SpazionIntornoAlleImmagini * 2;
                                AddImageToTheSlide(slide: slideToEdit,
                                             imageId: slideToGenerate.Contents[0],
                                             imageWidth: imageWidth,
                                             imageHeight: imageHeight,
                                             imagePostionY: offSetVerticale + SpazionIntornoAlleImmagini,
                                             imagePostionX: SpazionIntornoAlleImmagini);
                                AddImageToTheSlide(slide: slideToEdit,
                                            imageId: slideToGenerate.Contents[1],
                                            imageWidth: imageWidth,
                                            imageHeight: imageHeight,
                                            imagePostionY: offSetVerticale + SpazionIntornoAlleImmagini,
                                            imagePostionX: imageWidth + (SpazionIntornoAlleImmagini * 3));
                                break;

                            case 3:
                                //throw new NotImplementedException("Caso con 3 immagini per slide in orizontale non ancora gestito");
                                break;

                            default:
                                throw new ArgumentOutOfRangeException("Numero di immagini per slide non gestito");
                        }


                    }
                    if (slideToGenerate.LayoutType == LayoutTypes.Vertical)
                    {
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

        private IShape AddImageToTheSlide(ISlide slide, string imageId, decimal imageWidth, decimal imageHeight, decimal imagePostionY, decimal imagePostionX)
        {
            var imgFilePath = GetTmpFolderImagePathByImageId(Context.TmpFolder, imageId);
            var imgStream = new FileStream(imgFilePath, FileMode.Open, FileAccess.Read);
            slide.Shapes.AddPicture(imgStream);
            imgStream.Close();

            var shape = slide.Shapes[slide.Shapes.Count - 1];
            shape.Width = imageWidth;
            shape.Height = imageHeight;
            shape.Y = imagePostionY;
            shape.X = imagePostionX;

            return shape;
        }
    }
}