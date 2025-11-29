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
        readonly Microsoft.Office.Interop.Excel.Application excelApp = null;
        Workbook workbook = null;
        Worksheet worksheet = null;

        public ImageExtractor(string excelFilePath)
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false,
                DisplayAlerts = false,
                //ScreenUpdating = false, //non usare, disabilita la creazione delle immagini
                //Calculation = XlCalculation.xlCalculationManual
            };
            InteropServices_Helper.WaitForExcelReady(excelApp);

            InteropServices_Helper.PrioritizeExcelProcess();

            InteropServices_Helper.RetryComCall(() => workbook = excelApp.Workbooks.Open(excelFilePath));
            InteropServices_Helper.RetryComCall(() => workbook.RefreshAll());
            InteropServices_Helper.RetryComCall(() => excelApp.CalculateUntilAsyncQueriesDone());
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
                InteropServices_Helper.RetryComCall(() => range.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap));
               
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
                // Ignoro questa eccezione in quanto può capitare di tanto in tanto, quindi la ignoro in modo di poter tentare un altro tentativo
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
    }
}