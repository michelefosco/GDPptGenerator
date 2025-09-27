using FilesEditor.Helpers;
using System.Collections.Generic;

namespace FilesEditor.Entities
{
    internal class StepContext
    {
        public readonly Configurazione Configurazione;
        //
       // public readonly CreatePresentationsInput UpdateReportsInput;
        public readonly CreatePresentationsOutput UpdateReportsOutput;    // oggetto di output, qui vengono messe tutte le informazioni di output utili per interfaccia e controlli dei test
        //
        public FileDebugHelper DebugInfoLogger = new FileDebugHelper(null);
        //public InfoFileController InfoFileController;
        //public InfoFileReport InfoFileReport;
        //
        public Dictionary<string, object> Parameters = new Dictionary<string, object>();
        //
        public string ConfigurationFolder;
        public string ExcelDataSourceFile;

        public string OutputFolder;
        public string TmpFolder;
       // public string PowerPointOutputFile;

        public List<ItemToExport> ItemsToExportAsImage;
        //public List<SlideToGenerate> SlideToGenerateList;
        //
        public StepContext(CreatePresentationsInput updateReportsInput, Configurazione configurazione)
        {
            Configurazione = configurazione;
            //
          //  UpdateReportsInput = updateReportsInput;
            OutputFolder = updateReportsInput.OutputFolder;
            ConfigurationFolder = updateReportsInput.ConfigurationFolder;
            //
            UpdateReportsOutput = new CreatePresentationsOutput();
            UpdateReportsOutput.SettaConfigurazioneUsata(Configurazione);
        }
    }
}