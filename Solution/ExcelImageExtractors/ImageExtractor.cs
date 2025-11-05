using ExcelImageExtractors.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing.Imaging;
using System.Windows.Forms; // Per Clipboard


namespace ExcelImageExtractors
{
    public class ImageExtractor : IImageExtractor
    {
        readonly Microsoft.Office.Interop.Excel.Application excelApp = null;
        readonly Workbook workbook = null;
        Worksheet worksheet = null;

        public ImageExtractor(string excelFilePath)
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

                // Per funzionare questo codice deve essere eseguito nel thread principale dell'applicazione
                if (Clipboard.ContainsImage())
                {
                    // Recupera l'immagine dagli appunti
                    var clipboardImage = Clipboard.GetImage();
                    clipboardImage?.Save(destinationPath, ImageFormat.Png);
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