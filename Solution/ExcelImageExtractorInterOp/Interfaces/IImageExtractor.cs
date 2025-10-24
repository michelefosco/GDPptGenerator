namespace ExcelImageExtractorInterOp.Interfaces
{
    internal interface IImageExtractor
    {
        bool ExportImages(string workSheetName, string rangeAddress, string destinationPath);
        void Close();
    }
}
