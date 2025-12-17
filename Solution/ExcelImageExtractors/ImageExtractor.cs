using ExcelImageExtractors.Helpers;
using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms; // Per Clipboard


namespace ExcelImageExtractors
{
    public class ImageExtractor
    {
        private const string MESSAGGIO_ERRORE_COPY_PICTURE = "CopyPicture method of Range class failed";

        readonly Microsoft.Office.Interop.Excel.Application excelApp = null;
        Workbook workbook = null;
        Worksheet worksheet = null;

        public ImageExtractor(string excelFilePath, bool runRefreshAll = true)
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false,
                DisplayAlerts = false,
            };

            // Aspetto che Excel sia pronto
            //InteropServices_Helper.WaitForExcelReady(excelApp);

            // Attempt a harmless call to check if Excel is responsive
            int justToWaitForExcelReady = 0;
            InteropServices_Helper.RetryComCall(() => justToWaitForExcelReady = excelApp.Hwnd);


            // Aumento la priorità del processso Excel
            InteropServices_Helper.PrioritizeExcelProcess();

            // Apro il file Excel
            InteropServices_Helper.RetryComCall(() => workbook = excelApp.Workbooks.Open(excelFilePath));

            // Eseguo il RefreshAll se richiesto
            if (runRefreshAll)
            {
                InteropServices_Helper.RetryComCall(() => workbook.RefreshAll());
                InteropServices_Helper.RetryComCall(() => excelApp.CalculateUntilAsyncQueriesDone());
            }
        }

        public void ExportToImageFileOnFileSystem(string workSheetName, string rangeAddress, string destinationPath)
        {
            Range range = null;
            try
            {
                // seleziono il foglio
                InteropServices_Helper.RetryComCall(() => worksheet = workbook.Sheets[workSheetName]);

                // seleziono il range
                InteropServices_Helper.RetryComCall(() => range = worksheet.Range[rangeAddress]);

                // copio il range come immmagine nella Clipboard
                InteropServices_Helper.RetryComCall(
                                    action: () => range.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap),
                                    errorMessageToIgnore: MESSAGGIO_ERRORE_COPY_PICTURE);

                // salvo l'immagine dagli appunti sul file system
                SaveImageFromClipboardOnFile(destinationPath);
            }
            catch (Exception ex)
            {
                
                // Ignoro questa eccezione in quanto può capitare di tanto in tanto, quindi la ignoro in modo di poter tentare un altro tentativo
                //if (// !ex.Message.Equals("CopyPicture method of Range class failed"))
                if (!ex.Message.Equals(MESSAGGIO_ERRORE_COPY_PICTURE, StringComparison.Ordinal))
                {
                    throw ex;
                }
            }
            finally
            {
                if (range != null)
                { Marshal.ReleaseComObject(range); }

                if (worksheet != null)
                { Marshal.ReleaseComObject(worksheet); }
            }
        }


        public void Save()
        {
            InteropServices_Helper.RetryComCall(() => workbook.Save());
        }

        public void Close()
        {
            if (worksheet != null)
            {
                Marshal.ReleaseComObject(worksheet);
            }

            if (workbook != null)
            {
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
            }

            if (excelApp != null)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }

        private static void SaveImageFromClipboardOnFile(string destinationPath)
        {
            const int NUMERO_MASSIMO_TENTATIVI = 8;

            // A volta l'immagine non è immediatamente disponibile nella Clipboard, quindi asppeto qualche millisecondo e riprovo
            for (int attemptNumber = 1; attemptNumber <= NUMERO_MASSIMO_TENTATIVI; attemptNumber++)
            {
                // Per funzionare questo codice deve essere eseguito nel thread principale dell'applicazione
                if (Clipboard.ContainsImage())
                {
                    // Recupera l'immagine dagli appunti
                    var clipboardImage = Clipboard.GetImage();
                    if (clipboardImage != null)
                    {
                        clipboardImage.Save(destinationPath, ImageFormat.Png);
                        break;
                    }
                }
                Thread.Sleep(attemptNumber * 50);
            }
        }
    }
}