using ExcelImageExtractors.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms; // Per Clipboard


namespace ExcelImageExtractors
{
    public class ImageExtractor_Interop : IImageExtractor
    {
        readonly Microsoft.Office.Interop.Excel.Application excelApp = null;
        Workbook workbook = null;
        Worksheet worksheet = null;

        public ImageExtractor_Interop(string excelFilePath)
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false,
                DisplayAlerts = false
            };
            WaitForExcelReady(excelApp);

            RetryComCall(() => workbook = excelApp.Workbooks.Open(excelFilePath));

            RetryComCall(() => workbook.RefreshAll());

            RetryComCall(() => excelApp.CalculateUntilAsyncQueriesDone());

            //todo: decidere se salvare o no il file dopo i refresh, potrebbe essere utile ma durate molto
            //            workbook.Save();
        }



        public void TryToExportToImageFileOnFileSystem(string workSheetName, string rangeAddress, string destinationPath)
        {
            Range range = null;
            try
            {
                // seleziono il foglio
                worksheet = workbook.Sheets[workSheetName];

                // seleziono il range
                range = worksheet.Range[rangeAddress];

                // copio il range come immmagine nella Clipboard
                range.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);

                // A volta l'immagine non è immediatamente disponibile nella Clipboard, quindi asppeto qualche millisecondo e riprovo
                for (int attemptNumber = 1; attemptNumber <= 5; attemptNumber++)
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
                    else
                    {
                        Thread.Sleep(attemptNumber * 50);
                    }
                }
            }
            catch (Exception ex)
            {
                if (!ex.Message.Equals("CopyPicture method of Range class failed"))
                {
                    Close();

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





        private void WaitForExcelReady(Microsoft.Office.Interop.Excel.Application excelApp)
        {
            int maxWait = 200;  // max 20 seconds
            while (maxWait-- > 0)
            {
                try
                {
                    // Attempt a harmless call to check if Excel is responsive
                    var test = excelApp.Hwnd;
                    return; // no exception → Excel is ready
                }
                catch (COMException ex)
                {
                    // Excel is temporarily busy (edit mode, dialog, recalculating)
                    if ((uint)ex.ErrorCode == 0x80010001)
                    {
                        Thread.Sleep(100);
                        continue;
                    }

                    // If it's not RPC_E_CALL_REJECTED → rethrow
                    throw;
                }
            }

            throw new TimeoutException("Excel did not become ready in time.");
        }


        private void RetryComCall(System.Action action)
        {
            const uint RPC_E_CALL_REJECTED = 0x80010001;
            int retries = 15;

            while (true)
            {
                try
                {
                    action();
                    return;
                }
                catch (COMException ex) when ((uint)ex.ErrorCode == RPC_E_CALL_REJECTED)
                {
                    if (--retries == 0)
                        throw;

                    Thread.Sleep(100);
                }
            }
        }


        //private void WaitForExcelReady()
        //{
        //    while (excelApp.Ready == false || excelApp.CalculationState != XlCalculationState.xlDone)
        //    {
        //        Thread.Sleep(100);
        //    }
        //}


        //public void RefreshAllExcel(string filePath)
        //{
        //    var excelApp = new Microsoft.Office.Interop.Excel.Application();
        //    excelApp.Visible = false;
        //    excelApp.DisplayAlerts = false;

        //    Microsoft.Office.Interop.Excel.Workbook wb = excelApp.Workbooks.Open(filePath);

        //    // Refresh All (connessioni, tabelle, PQ, Pivot, tutto)
        //    wb.RefreshAll();

        //    // Attendi che Excel finisca i refresh
        //    excelApp.CalculateUntilAsyncQueriesDone();

        //    // Salva e chiudi
        //    wb.Save();
        //    wb.Close(false);
        //    excelApp.Quit();

        //    // Rilascio COM (importantissimo)
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        //}
    }
}