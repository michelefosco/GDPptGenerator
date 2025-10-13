using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using System.Collections.Generic;

namespace FilesEditor.Entities
{
    public class StepContext: UserInterfaceInputBase
    {
        public readonly Configurazione Configurazione;
        public FileDebugHelper DebugInfoLogger = new FileDebugHelper(null);
        //
//        public Dictionary<string, object> Parameters = new Dictionary<string, object>();
        //
        public string OutputDataSourceFilePath;


        public EsitiFinali Esito { get; private set; }

        // public string PowerPointOutputFile;
        public List<ItemToExport> ItemsToExportAsImage;
        //public List<SlideToGenerate> SlideToGenerateList;
        //

        public StepContext(Configurazione configurazione)
        {
            Configurazione = configurazione;
        }

        public void SettaEsitoFinale(EsitiFinali esito)
        {
            Esito = esito;
        }
        public void SetContextFromInput(BuildPresentationInput input)
        {
            if (input == null) { return; }
            base.DestinationFolder = input.DestinationFolder;
            base.TmpFolder = input.TmpFolder;
            base.SourceFilesFolder = input.SourceFilesFolder;
            base.FileDebugPath = input.FileDebugPath;
        }

        public void SetContextFromInput(ValidaSourceFilesInput input)
        {
            if (input == null) { return; }
            base.DestinationFolder = input.DestinationFolder;
            base.TmpFolder = input.TmpFolder;
            base.SourceFilesFolder = input.SourceFilesFolder;
            base.FileDebugPath = input.FileDebugPath;
        }
    }
}