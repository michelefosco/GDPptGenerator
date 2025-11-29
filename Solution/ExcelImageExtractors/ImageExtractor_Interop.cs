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
                DisplayAlerts = false,
                ScreenUpdating = false,
                //Calculation = XlCalculation.xlCalculationManual
            };
            ExcelInterops_Helper.WaitForExcelReady(excelApp);

            ExcelInterops_Helper.PrioritizeExcelProcess();

            ExcelInterops_Helper.RetryComCall(() => workbook = excelApp.Workbooks.Open(excelFilePath));
            ExcelInterops_Helper.RetryComCall(() => workbook.RefreshAll());
            ExcelInterops_Helper.RetryComCall(() => excelApp.CalculateUntilAsyncQueriesDone());

            //todo: decidere se salvare o no il file dopo i refresh, potrebbe essere utile ma durate molto
            ExcelInterops_Helper.RetryComCall(() => workbook.Save());
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