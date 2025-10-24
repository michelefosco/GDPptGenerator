using ExcelImageExtractorInterOp.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms; // Per Clipboard


namespace ImageExtractor
{
    public class ImageExtractor : IImageExtractor
    {
        readonly Microsoft.Office.Interop.Excel.Application excelApp = null;
        readonly Workbook workbook = null;
        Worksheet worksheet = null;

        public ImageExtractor(string excelFilePath)
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            workbook = excelApp.Workbooks.Open(excelFilePath);
        }

        public bool ExportImages(string workSheetName, string rangeAddress, string destinationPath)
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
                // Recupera l'immagine dagli appunti
                Image clipboardImage = null;
                if (Clipboard.ContainsImage())
                {
                    clipboardImage = Clipboard.GetImage();
                    if (clipboardImage != null)
                    { clipboardImage.Save(destinationPath, ImageFormat.Png); }
                }

                // Segnalo se l'estrazione dell'immagine non è andata a buon fine
                if (clipboardImage == null)
                { throw new Exception($"la Clipboard non contiene un'immagine come previsto per il file '{destinationPath}'"); }

                return true;
            }
            catch (Exception ex)
            {
                //todo rimuovere dopo i test
                var a = ex.Message;
                var b = ex.InnerException;
                return false;
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