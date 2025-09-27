using FilesEditor.Constants;
using FilesEditor.Entities;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FilesEditor.Steps
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

            var powerPointStructFiles = GetFilesListFromFolder(Context.ConfigurationFolder, FiltriNomifile.FILTRO_NOME_FILE_PPT_STRUTTURA);

            foreach (var percorsoFile in powerPointStructFiles)
            {
                var slideToGenerateList = getListaSlidesDaFile(percorsoFile);

                var suffisso = Path.GetFileNameWithoutExtension(percorsoFile).Replace("PowerPoint_Struttura", "");

                var outputfilePath = $"{Context.OutputFolder}\\Output{suffisso}.pptx";
                // ripulisco il possibile file di output
                if (File.Exists(outputfilePath))
                { File.Delete(outputfilePath); }


                // copio il template nella cartella di output
                var percorsoFileTemplatePowerPoint = Path.Combine(Context.ConfigurationFolder, Constants.FileNames.POWERPOINT_TEMPLATE_FILENAME);
                File.Copy(percorsoFileTemplatePowerPoint, outputfilePath);
                var pres = new ShapeCrawler.Presentation(outputfilePath);


                // predispongo il numero di slides necessarie
                //var numeroSlidesNecessarieAggiungere = slideToGenerateList.Count - 1; // tolgo 1 perchè c'è già una slide template
                //for (int j = 1; j <= numeroSlidesNecessarieAggiungere; j++)
                //{
                //    pres.Slides.Add(SLIDE_TEMPLATE_INDEX, pres.Slides.Count + 1);
                //}

                const int SpazionIntornoAlleImmagini = 10;

                int slideToEditIndex = 3; // la slide 1 è il titolo, la 2 è l'indice, la 3 è la prima slide template
                foreach (var slideToGenerate in slideToGenerateList)
                {
                    // duplico la slide template
                    pres.Slides.Add(pres.Slide(SLIDE_TEMPLATE_1_INDEX), pres.Slides.Count + 1);
                    var slideToEdit = pres.Slide(pres.Slides.Count);

                    var titleTextBox = slideToEdit.GetTextBoxes().FirstOrDefault(tb => tb.Text.Contains("Titolo"));
                    if (titleTextBox != null)
                    {
                        titleTextBox.SetText(slideToGenerate.Title);
                    }


                    const int offSetVerticale = 80;
                    // in base al template inserisco una o più immagini in diverse posizioni
                    string imgFilePath;// = GetImagePath(slideToGenerate.ImageId1);
                    FileStream imgStream;// = new FileStream(imgFilePath, FileMode.Open, FileAccess.Read);
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
                            altezzaImmagine = (pres.SlideHeight - offSetVerticale) / numeroImmaginiInVerticale - SpazionIntornoAlleImmagini * 2;
                            //
                            imgFilePath = GetImagePath(slideToGenerate.ImageId1);
                            imgStream = new FileStream(imgFilePath, FileMode.Open, FileAccess.Read);
                            slideToEdit.Shapes.AddPicture(imgStream);
                            slideToEdit.Shapes[slideToEdit.Shapes.Count - 1].Y = offSetVerticale + SpazionIntornoAlleImmagini;
                            slideToEdit.Shapes[slideToEdit.Shapes.Count - 1].X = SpazionIntornoAlleImmagini;
                            slideToEdit.Shapes[slideToEdit.Shapes.Count - 1].Width = larghezzaImmagine;
                            slideToEdit.Shapes[slideToEdit.Shapes.Count - 1].Height = altezzaImmagine;
                            imgStream.Close();
                            break;
                        case 2:
                            // 2 immagini sulla stessa riga
                            numeroImmaginiInVerticale = 1;
                            numeroImmaginiInOrizzontale = 2;
                            larghezzaImmagine = pres.SlideWidth / numeroImmaginiInOrizzontale - SpazionIntornoAlleImmagini * 2;
                            altezzaImmagine = (pres.SlideHeight - offSetVerticale) / numeroImmaginiInVerticale - SpazionIntornoAlleImmagini * 2;
                            //
                            imgFilePath = GetImagePath(slideToGenerate.ImageId1);
                            imgStream = new FileStream(imgFilePath, FileMode.Open, FileAccess.Read);
                            slideToEdit.Shapes.AddPicture(imgStream);
                            slideToEdit.Shapes[slideToEdit.Shapes.Count - 1].Y = offSetVerticale + SpazionIntornoAlleImmagini;
                            slideToEdit.Shapes[slideToEdit.Shapes.Count - 1].X = SpazionIntornoAlleImmagini;
                            slideToEdit.Shapes[slideToEdit.Shapes.Count - 1].Width = larghezzaImmagine;
                            slideToEdit.Shapes[slideToEdit.Shapes.Count - 1].Height = altezzaImmagine;
                            imgStream.Close();
                            //
                            imgFilePath = GetImagePath(slideToGenerate.ImageId2);
                            imgStream = new FileStream(imgFilePath, FileMode.Open, FileAccess.Read);
                            slideToEdit.Shapes.AddPicture(imgStream);
                            slideToEdit.Shapes[slideToEdit.Shapes.Count - 1].Y = offSetVerticale + SpazionIntornoAlleImmagini;
                            slideToEdit.Shapes[slideToEdit.Shapes.Count - 1].X = larghezzaImmagine + (SpazionIntornoAlleImmagini * 3);
                            slideToEdit.Shapes[slideToEdit.Shapes.Count - 1].Width = larghezzaImmagine;
                            slideToEdit.Shapes[slideToEdit.Shapes.Count - 1].Height = altezzaImmagine;
                            imgStream.Close();
                            //
                            break;
                        case 3:

                            break;
                        case 4:

                            break;
                        default:
                            break;
                    }

                    slideToEditIndex++;
                }











                #region Operazioni finali
                // rimuovo i 4 templale
                for (var slideIndex = 1; slideIndex <= NumberOfTemplateSlides; slideIndex++)
                { pres.Slide(SLIDE_TEMPLATE_1_INDEX).Remove(); }

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
                var slideType = int.Parse(campi[0].Trim());
                var title = campi[1].Trim();
                var imageId1 = campi[2].Trim().ToUpper();
                var imageId2 = campi[3].Trim().ToUpper();
                var imageId3 = campi[4].Trim().ToLower();

                SlideToGenerateList.Add(new SlideToGenerate
                {
                    Title = title,
                    SlideType = slideType,
                    ImageId1 = imageId1,
                    ImageId2 = imageId2,
                    ImageId3 = imageId3
                });
            }

            return SlideToGenerateList;
        }
    }
}