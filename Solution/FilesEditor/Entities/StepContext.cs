using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Helpers;
using System.Collections.Generic;

namespace FilesEditor.Entities
{
    internal class StepContext
    {
        public readonly Configurazione Configurazione;
        //
        public readonly CreatePresentationsInput CreatePresentationsInput;
        public readonly CreatePresentationsOutput CreatePresentationsOutput;    // oggetto di output, qui vengono messe tutte le informazioni di output utili per interfaccia e controlli dei test
        //
        public FileDebugHelper DebugInfoLogger = new FileDebugHelper(null);
        //public InfoFileController InfoFileController;
        //public InfoFileReport InfoFileReport;
        //
        public Dictionary<string, object> Parameters = new Dictionary<string, object>();
        //
        public string TemplatesFolder;
        public string ExcelDataSourceFile;

        public string OutputFolder;
        public string TmpFolder;
        // public string PowerPointOutputFile;

        public List<ItemToExport> ItemsToExportAsImage;
        //public List<SlideToGenerate> SlideToGenerateList;
        //
        public StepContext(CreatePresentationsInput createPresentationsInput, Configurazione configurazione)
        {
            Configurazione = configurazione;
            //
            CreatePresentationsInput = createPresentationsInput;
            //
            OutputFolder = createPresentationsInput.OutputFolder;
            TmpFolder = createPresentationsInput.TmpFolder;
            TemplatesFolder = createPresentationsInput.TemplatesFolder;
            //
            CreatePresentationsOutput = new CreatePresentationsOutput();
        }
    }
}