namespace FilesEditor.Entities
{
    public class ItemToExport
    {
        public readonly string WorkSheetName;
        public readonly string PrintArea;
        public readonly string ImageId;
        public readonly string ImageFilePath;

        public bool IsPresentOnFileSistem { get; private set; }

        public ItemToExport(string workSheetName,
                            string printArea,
                            string imageId,
                            string imageFilePath)
        {
            WorkSheetName = workSheetName;
            PrintArea = printArea;
            ImageId = imageId;
            ImageFilePath = imageFilePath;
            IsPresentOnFileSistem = false;
        }

        public void MarkAsPresentOnFileSistem()
        {
            IsPresentOnFileSistem = true;
        }
    }
}