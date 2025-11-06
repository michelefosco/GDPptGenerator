using ExcelImageExtractors.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing.Imaging;
using System.Threading;
using System.Windows.Forms; // Per Clipboard


namespace ExcelImageExtractors
{
    public class ImageExtractor_Interop : IImageExtractor
    {
        readonly Microsoft.Office.Interop.Excel.Application excelApp = null;
        readonly Workbook workbook = null;
        Worksheet worksheet = null;

        public ImageExtractor_Interop(string excelFilePath)
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false
            };
            workbook = excelApp.Workbooks.Open(excelFilePath);
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

                // A volta l'immagine non è immediatamente disponibile nella Clipboard
                // per questo eseguo più tentantivi
                for (int i = 1; i <= 5; i++)
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
                        Thread.Sleep(100);
                    }
                }
            }
            catch (Exception ex)
            {
                if (!ex.Message.Equals("CopyPicture method of Range class failed"))
                {
                    throw ex;
                }
            }
            finally
            {
                if (range != null)
                { System.Runtime.InteropServices.Marshal.ReleaseComObject(range); }

                if (worksheet != null)
                { System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet); }
            }
        }

        public void Close()
        {
            if (worksheet != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            }

            if (workbook != null)
            {
                workbook.Close(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            }

            if (excelApp != null)
            {
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}