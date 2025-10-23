using Microsoft.Office.Interop.Excel;
using System.Drawing.Imaging;
using System.Windows.Forms; // Add reference to System.Windows.Forms
using System.Threading;


namespace ImagesFromExcelGenerator
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
           var excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            var workbook = excelApp.Workbooks.Open(_excelFilePath);
            var sheet = workbook.Sheets[workSheetName];
            var range = sheet.get_Range(rangeAddress);

            range.CopyPicture(Microsoft.Office.Interop.Excel.XlPictureAppearance.xlScreen, Microsoft.Office.Interop.Excel.XlCopyPictureFormat.xlBitmap);

            // Clipboard access requires STA thread, use a new thread
            Thread thread = new Thread(() =>
            {
                if (Clipboard.ContainsImage())
                {
                    var img = Clipboard.GetImage();
                    img.Save(destinationPath, ImageFormat.Png);
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            workbook.Close(false);
            excelApp.Quit();
        }
    }

    // Usage Example:
    // ExcelImageSaver.SaveRangeAsImage(@"C:\path\file.xlsx", "Sheet1", "A1:C10", @"C:\output\image.png");

}