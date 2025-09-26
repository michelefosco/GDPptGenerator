//
using System.Collections.Generic;


namespace POC_Console.Entities
{
    public class Context
    {
        public string ConfigurationFolder;
        public string ExcelDataSourceFile;

        public string OutputFolder;
        public string TmpFolder;
        public string PowerPointOutputFile;


        public List<ItemToExport> ItemsToExportAsImage;
        //public List<SlideToGenerate> SlideToGenerateList;
    }
}