using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using System.Collections.Generic;

namespace FilesEditor.Entities
{
    public class StepContext : UserInterfaceInputBase
    {
        public readonly Configurazione Configurazione;
        public DebugInfoLogger DebugInfoLogger = new DebugInfoLogger(null);
        public List<string> Warnings = new List<string>();

        //
        //        public Dictionary<string, object> Parameters = new Dictionary<string, object>();
        //
        public string OutputDataSourceFilePath;
        public List<InputDataFilters_Items> Applicablefilters = new List<InputDataFilters_Items>();
        public List<SlideToGenerate> SildeToGenerate = new List<SlideToGenerate>();

        public EsitiFinali Esito { get; private set; }

        // public string PowerPointOutputFile;
        public List<ItemToExport> ItemsToExportAsImage;
        //public List<SlideToGenerate> SlideToGenerateList;

        // Specifici di BuildPresentation
        //   bool replaceAllData_FileSuperDettagli,
        //DateTime periodDate


        // Specifici di ValidaSourceFilesInput
        public string FileBudgetPath { get; private set; }
        public string FileForecastPath { get; private set; }
        public string FileSuperDettagliPath { get; private set; }
        public string FileRunRatePath { get; private set; }

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
            //
            FileBudgetPath = input.FileBudgetPath;
            FileForecastPath = input.FileForecastPath;
            FileSuperDettagliPath = input.FileSuperDettagliPath;
            FileRunRatePath = input.FileRunRatePath;
        }
    }
}