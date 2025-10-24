using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Threading;
using System.Windows.Forms; // Per Clipboard


namespace ExcelImageExtractorInterOp
{
    public class ExcelImageSaver
    {

        string _excelFilePath;

        public ExcelImageSaver(string excelFilePath)
        {
            _excelFilePath = excelFilePath;
        }

        public void Close()
        {
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }

        public void ExportImages(string workSheetName, string rangeAddress, string destinationPath)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            Range range = null;
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false;
                workbook = excelApp.Workbooks.Open(_excelFilePath);
                worksheet = workbook.Sheets[workSheetName];

                //  string rangeAddress = $"{startRange}:{endRange}";
                range = worksheet.Range[rangeAddress];

                range.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);

                // Recupera l'immagine dagli appunti
                // Per funzionare questo codice non deve essere eseguito in un thread separato 
                Image clipboardImage = null;
                if (Clipboard.ContainsImage())
                {
                    clipboardImage = Clipboard.GetImage();
                }
                else
                {
                    throw new Exception($"The clipboard non contiene un'immagine come previsto per il file '{destinationPath}'");
                }

                if (clipboardImage != null)
                {
                    clipboardImage.Save(destinationPath, ImageFormat.Png);
                }


                //Ok con UnitTest no con Windows
                //if (System.Windows.Forms.Clipboard.ContainsImage())
                //{
                //    var img = System.Windows.Forms.Clipboard.GetImage();
                //    img.Save(destinationPath, ImageFormat.Png);
                //}
                //else
                //{
                //    throw new Exception($"Clipboard vuota per il file immagine da generare '{destinationPath}'");
                //}

                #region OK, funziona tramite gli unit tests
                // Clipboard access requires STA thread, use a new thread
                //Thread thread = new Thread(() =>
                //{
                //    if (System.Windows.Forms.Clipboard.ContainsImage())
                //    {
                //        var img = System.Windows.Forms.Clipboard.GetImage();
                //        img.Save(destinationPath, ImageFormat.Png);
                //    }
                //    else
                //    {
                //        throw new Exception($"Clipboard vuota per il file immagine da generare '{destinationPath}'");
                //    }
                //});

                //thread.SetApartmentState(ApartmentState.STA);
                //thread.Start();
                //thread.Join();
                #endregion

                workbook.Close(false);
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception ex)
            {
                var a = ex.Message;
                var b = ex.InnerException;
            }
            finally
            {
                if (range != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }

        }
    }
}