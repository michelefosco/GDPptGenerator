using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using System.Collections.Generic;

namespace FilesEditor.Entities
{
    internal class StepContext
    {
        public readonly Configurazione Configurazione;
        //
        public readonly BuildPresentationInput BuildPresentationInput;
        public readonly BuildPresentationOutput BuildPresentationOutput;    // oggetto di output, qui vengono messe tutte le informazioni di output utili per interfaccia e controlli dei test
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
        public StepContext(BuildPresentationInput buildPresentationInput, Configurazione configurazione)
        {
            Configurazione = configurazione;
            //
            BuildPresentationInput = buildPresentationInput;
            //
            OutputFolder = buildPresentationInput.OutputFolder;
            TmpFolder = buildPresentationInput.TmpFolder;
            TemplatesFolder = buildPresentationInput.TemplatesFolder;
            //
            BuildPresentationOutput = new BuildPresentationOutput(EsitiFinali.Undefined);
        }
    }
}