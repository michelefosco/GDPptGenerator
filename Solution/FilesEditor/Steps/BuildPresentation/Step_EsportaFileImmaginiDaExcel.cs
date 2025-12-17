using FilesEditor.Entities;
using FilesEditor.Enums;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
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
        internal override string StepName => "Step_EsportaFileImmaginiDaExcel";

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
      //      throw new Exception("Fake exception to test retries logic.");

            ThrowExceptionIn_AnImageAlreadyExists();

            // monitora se il salvataggio del file Excel è già stato fatto almeno una volta
            var alreadySavedOnce = false;

            int attemptNumber = 0;

            const int MAX_NUMBER_OF_fullAttempts = 5;
            const int MAX_NUMBER_OF_imageExtractionAttempts = 10;

            // eseguo più tentativi completi di apertura file ed estrazione immagini
            for (int fullAttemptNumber = 1; fullAttemptNumber <= MAX_NUMBER_OF_fullAttempts; fullAttemptNumber++)
            {
                var startTime = DateTime.UtcNow;
                // alla prima esecuzione apro il file Excel e faccio il "Refresh All" dei dati collegati
                var imageExtractor = new ExcelImageExtractors.ImageExtractor(Context.DataSourceFilePath, !alreadySavedOnce);  // Interop
                Context.DebugInfoLogger.LogPerformance(StepName + $" Istanziamento ImageExtractor (apertura Excel e refresh data)", DateTime.UtcNow - startTime);

                // eseguo più tentativi di estrazione immagini
                for (int imageExtractionAttemptNumber = 1; imageExtractionAttemptNumber <= MAX_NUMBER_OF_imageExtractionAttempts; imageExtractionAttemptNumber++)
                {
                    attemptNumber++;

                    // Processo tutti gli elementi non ancora presenti sul file system
                    foreach (var itemToExportAsImage in Context.ItemsToExportAsImage.Where(_ => !_.IsPresentOnFileSistem))
                    {
                        startTime = DateTime.UtcNow;
                        // tento di generare il file, alcune volte potrebbe non funzionare al primo tentativo
                        imageExtractor.ExportToImageFileOnFileSystem(itemToExportAsImage.WorkSheetName, itemToExportAsImage.PrintArea, itemToExportAsImage.ImageFilePath);
                        var timeSpent = DateTime.UtcNow - startTime;

                        // verifico la sua presenza su file system ed eventualmente lo marco come "Presente"
                        if (File.Exists(itemToExportAsImage.ImageFilePath))
                        { itemToExportAsImage.MarkAsPresentOnFileSistem(); }

                        // loggo il risultato del tentativo
                        Context.DebugInfoLogger.LogPerformance(StepName + $" ExportToImageFileOnFileSystem {itemToExportAsImage.WorkSheetName}  {itemToExportAsImage.PrintArea}", timeSpent);
                        Context.DebugInfoLogger.LogInfoExportImages(
                            attemptNumber, //imageExtractionAttemptNumber,
                            itemToExportAsImage.ImageId,
                            itemToExportAsImage.ImageFilePath,
                            itemToExportAsImage.WorkSheetName,
                            itemToExportAsImage.PrintArea,
                            itemToExportAsImage.IsPresentOnFileSistem,
                            timeSpent
                            );
                    }

                    // se tutti i file sono presenti interrompo i tentativi 
                    if (!Context.ItemsToExportAsImage.Any(_ => !_.IsPresentOnFileSistem))
                    { break; }

                    Thread.Sleep(500); // aspetto mezzo secondo prima di un nuovo tentativo
                }

                // salvo le modifiche al file Excel solo la prima volta (in quanto modificao dal RefreshAll durante il caricamento)
                if (!alreadySavedOnce)
                {
                    startTime = DateTime.UtcNow;
                    imageExtractor.Save();
                    Context.DebugInfoLogger.LogPerformance(StepName + $" imageExtractor.Save()", DateTime.UtcNow - startTime);

                    Context.SetDatasourceStatus_RefreshAllCompletato();
                    alreadySavedOnce = true;
                }

                // rilascio le risorse e chiudo il processo di Excel
                imageExtractor.Close();

                // se tutti i file sono presenti interrompo i tentativi 
                if (!Context.ItemsToExportAsImage.Any(_ => !_.IsPresentOnFileSistem))
                { break; }

                Thread.Sleep(500); // aspetto mezzo secondo prima di un nuovo tentativo
            }

            ThrowExceptionIf_AnImageIsMissing();
        }

        private void ThrowExceptionIn_AnImageAlreadyExists()
        {
            foreach (var itemsToExportAsImage in Context.ItemsToExportAsImage)
            {
                if (File.Exists(itemsToExportAsImage.ImageFilePath))
                {
                    throw new System.Exception($"The file '{itemsToExportAsImage.ImageFilePath}' should not be there yet!");
                }
            }
        }

        private void ThrowExceptionIf_AnImageIsMissing()
        {
            foreach (var itemsToExportAsImage in Context.ItemsToExportAsImage)
            {
                if (!File.Exists(itemsToExportAsImage.ImageFilePath))
                {
                    throw new System.Exception($"The data source was updated successfully, but the presentation could not be built.\r\nOne required image was not extracted successfully.\r\nPlease click Build Presentation Only to try again.");
                }
            }
        }
    }
}