using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using System;
using System.Collections.Generic;

namespace FilesEditor.Entities
{
    public class StepContext : UserInterfaceInputBase
    {
        public EsitiFinali Esito { get; private set; }
        public Configurazione Configurazione;
        public DebugInfoLogger DebugInfoLogger = new DebugInfoLogger(null);
        public List<string> Warnings = new List<string>();
        //        public Dictionary<string, object> Parameters = new Dictionary<string, object>();
        public string OutputDataSourceFilePath;
        public List<InputDataFilters_Item> Applicablefilters = new List<InputDataFilters_Item>();
        public List<SlideToGenerate> SildeToGenerate = new List<SlideToGenerate>();
        public bool ReplaceAllData_FileSuperDettagli { get; private set; }
        public DateTime PeriodDate { get; private set; }

        public List<ItemToExport> ItemsToExportAsImage = new List<ItemToExport>();
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
            //
            ReplaceAllData_FileSuperDettagli = input.ReplaceAllData_FileSuperDettagli;
            PeriodDate = input.PeriodDate;
            Applicablefilters = input.Applicablefilters ?? new List<InputDataFilters_Item>();
        }

        public void SetContextFromInput(ValidateSourceFilesInput input)
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