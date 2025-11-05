using ExcelImageExtractors.Interfaces;
using FilesEditor.Entities;
using FilesEditor.Enums;
using System;
using System.IO;
using System.Linq;
using System.Threading;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_EsportaFileImmaginiDaExcel : StepBase
    {
        public override string StepName => "Step_EsportaFileImmaginiDaExcel";

        internal override void BeforeTask()
        {
            Context.DebugInfoLogger.LogStepContext(StepName, Context);
        }

        internal override void ManageInfoAboutPerformedStepTask(TimeSpan timeSpent)
        {
            Context.DebugInfoLogger.LogPerformance(StepName, timeSpent);
        }

        internal override void AfterTask()
        {
            Context.DebugInfoLogger.LogStepContext(StepName, Context);
        }
        public Step_EsportaFileImmaginiDaExcel(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            CreaFilesImmaginiDaEsportare();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        /// <summary>
        /// Genera le immagini da usarsi nelle slides
        /// </summary>
        private void CreaFilesImmaginiDaEsportare()
        {
            #region Verifico che i file delle immagini da generare per le slide non esistano già
            foreach (var itemsToExportAsImage in Context.ItemsToExportAsImage)
            {
                if (File.Exists(itemsToExportAsImage.ImageFilePath))
                {
                    throw new System.Exception($"The file '{itemsToExportAsImage.ImageFilePath}' should not be there yet!");
                }
            }
            #endregion

            // todo: predisposta la possibilità di usare Aspose.Cell in casi estremi.
            // come settare questa opzione?
            var useIterops = true;
            int millisecondsWaitingTimeBetweenExtractions = 0;
            const int maxNumberOfAttempts = 15;

            int numberOfFailures = 0;

            for (int attemptNumber = 1; attemptNumber <= maxNumberOfAttempts; attemptNumber++)
            {
                var imageExtractor = (useIterops)
                    ? (IImageExtractor)new ExcelImageExtractors.ImageExtractor(Context.DataSourceFilePath)
                    : (IImageExtractor)new ExcelImageExtractors.ImageExtractor_Aspose(Context.DataSourceFilePath);  // Aspose.Cell


                // Processo tutti gli elementi non ancora presenti sul file system
                foreach (var itemToExportAsImage in Context.ItemsToExportAsImage.Where(_ => !_.IsPresentOnFileSistem))
                {
                    // tento di generare il file, alcune volte potrebbe non funzionare al primo tentativo
                    imageExtractor.TryToExportToImageFileOnFileSystem(itemToExportAsImage.WorkSheetName, itemToExportAsImage.PrintArea, itemToExportAsImage.ImageFilePath);

                    // Tempo d'attesa tra un'immagine e la successiva per agevolare il rilascio delle risorse, chiusura Excel, etc.
                    Thread.Sleep(millisecondsWaitingTimeBetweenExtractions);

                    // verifico la sua presenza su file system ed eventualmente lo marco come "Presente"
                    if (File.Exists(itemToExportAsImage.ImageFilePath))
                    { 
                        itemToExportAsImage.MarkAsPresentOnFileSistem(); 
                    }
                    else
                    {
                        numberOfFailures++;
                    }

                    Context.DebugInfoLogger.LogInfoExportImages(
                        attemptNumber,
                        itemToExportAsImage.ImageId,
                        itemToExportAsImage.ImageFilePath,
                        itemToExportAsImage.WorkSheetName,
                        itemToExportAsImage.PrintArea,
                        itemToExportAsImage.IsPresentOnFileSistem
                        );
                }


                // rilascio le risorse e chiudo il processo di Excel
                imageExtractor.Close();

                // se tutti i file sono presenti interrompo i tentativi 
                if (!Context.ItemsToExportAsImage.Any(_ => !_.IsPresentOnFileSistem))
                {
                    break;
                }
                else
                {
                    // allungo il tempo d'attesa tra un'immagine e la successiva
                    millisecondsWaitingTimeBetweenExtractions += 0;
                }
            }
        }
    }
}