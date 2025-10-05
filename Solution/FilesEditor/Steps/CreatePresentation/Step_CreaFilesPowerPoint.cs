using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using ShapeCrawler;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FilesEditor.Steps.CreatePresentation
{
    internal class Step_CreaFilesPowerPoint : Step_Base
    {
        public Step_CreaFilesPowerPoint(StepContext context) : base(context)
        { }

        internal override CreatePresentationsOutput DoSpecificTask()
        {
            creazionePowerPoint();
            return null; // Step intermedio, non ritorna alcun esito
        }

        const int NumberOfTemplateSlides = 1;
        const int SLIDE_TEMPLATE_1_INDEX = 3;
        private void creazionePowerPoint()
        {
            var powerPointStructFiles = GetFilesListFromFolder(Context.TemplatesFolder, FileNames.FILTRO_NOME_FILE_PPT_STRUTTURA);

            foreach (var percorsoFile in powerPointStructFiles)
            {
                var slideToGenerateList = getListaSlidesDaFile(percorsoFile);


                #region Predispongo il file di output
                var suffisso = Path.GetFileNameWithoutExtension(percorsoFile).Replace("PowerPoint_Struttura", "");
                var outputfilePath = $"{Context.OutputFolder}\\Output{suffisso}.pptx";
                // ripulisco il possibile file di output
                if (File.Exists(outputfilePath))
                { File.Delete(outputfilePath); }
                #endregion


                #region Copio il template nella cartella di output
                var percorsoFileTemplatePowerPoint = Path.Combine(Context.TemplatesFolder, Constants.FileNames.POWERPOINT_TEMPLATE_FILENAME);
                File.Copy(percorsoFileTemplatePowerPoint, outputfilePath);
                var pres = new ShapeCrawler.Presentation(outputfilePath);
                #endregion


                const int SpazionIntornoAlleImmagini = 10;
                const int offSetVerticale = 80;


                for (int j = 1; j <= slideToGenerateList.Count - 1; j++)
                {
                    pres.Slides.Add(pres.Slide(SLIDE_TEMPLATE_1_INDEX), pres.Slides.Count + 1);
                }

                int slideToEditIndex = SLIDE_TEMPLATE_1_INDEX;
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
                                throw new NotImplementedException("Caso con 3 immagini per slide in orizontale non ancora gestito");
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
                #endregion
            }
        }


        //todo read from Excel
        private List<SlideToGenerate> getListaSlidesDaFile(string percorsoFile)
        {
            var SlideToGenerateList = new List<SlideToGenerate>();

            // Legge tutte le righe del file
            string[] righe = File.ReadAllLines(percorsoFile);

            foreach (string riga in righe)
            {
                // Divide la riga nei campi separati da ";"
                string[] campi = riga.Split(';');



                //todo: ragionare su queste trasformazioni
                var slideType = (LayoutTypes) int.Parse(campi[0].Trim());
                var title = campi[1].Trim();                             

                var contents = new List<string>();
                var content1 = campi[2].Trim().ToUpper();
                if (!string.IsNullOrWhiteSpace(content1)) { contents.Add(content1); }

                var content2 = campi[3].Trim().ToUpper();
                if (!string.IsNullOrWhiteSpace(content1)) { contents.Add(content2); }

                var content3 = campi[4].Trim().ToLower();
                if (!string.IsNullOrWhiteSpace(content1)) { contents.Add(content3); }

                SlideToGenerateList.Add(new SlideToGenerate
                {
                    OutputFileName = "Presentazione.pptx", //todo: read from excel
                    Title = title,
                    LayoutType = slideType,
                    Contents = contents           
                });
            }

            return SlideToGenerateList;
        }

        private IShape AddImageToTheSlide(ISlide slide, string imageId, decimal imageWidth, decimal imageHeight, decimal imagePostionY, decimal imagePostionX)
        {
            var imgFilePath = GetImagePath(imageId);
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