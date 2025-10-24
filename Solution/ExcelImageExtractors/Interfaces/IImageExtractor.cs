namespace ExcelImageExtractors.Interfaces
{
    public interface IImageExtractor
    {
        void TryToExportToImageFileOnFileSystem(string workSheetName, string rangeAddress, string destinationPath);
        void Close();
    }
}
