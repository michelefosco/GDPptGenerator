namespace FilesEditor.Entities.MethodsArgs
{
    public class UserInterfaceInputBase
    {
        public string DestinationFolder { get; set; }
        public string TmpFolder { get; set; }
        public string DataSourceFilePath { get; set; }
        public string DebugFilePath { get; set; }


        //
        public string FileBudgetPath { get; set; }
        public string FileForecastPath { get; set; }
        public string FileSuperDettagliPath { get; set; }
        public string FileRunRatePath { get; set; }
    }
}