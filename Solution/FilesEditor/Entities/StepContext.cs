using EPPlusExtensions;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using System;
using System.Collections.Generic;

namespace FilesEditor.Entities
{
    internal class StepContext : UserInterfaceInputBase
    {
        // Input
        public bool ReplaceAllData_FileSuperDettagli { get; private set; }
        public DateTime PeriodDate { get; private set; }

        // Base class already has these properties
        // DestinationFolder
        // TmpFolder
        // DataSourceFilePath
        // DebugFilePath
        // FileBudgetPath
        // FileForecastPath
        // FileSuperDettagliPath
        // FileRunRatePath

        //        public Dictionary<string, object> Parameters = new Dictionary<string, object>();


        public Configurazione Configurazione;

        public EsitiFinali Esito { get; private set; }
        


        private EPPlusHelper _ePPlusHelperDataSource;
        public EPPlusHelper ePPlusHelperDataSource
        {
            get
            {
                if (_ePPlusHelperDataSource == null)
                {
                    if (string.IsNullOrEmpty(DataSourceFilePath))
                    { throw new Exception("Inizializzare 'DataSourceFilePath' prima di usare 'ePPlusHelperDataSource'"); }

                    _ePPlusHelperDataSource = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(DataSourceFilePath, FileTypes.DataSource);
                }
                return _ePPlusHelperDataSource;
            }
        }
        
        public DebugInfoLogger DebugInfoLogger = new DebugInfoLogger(null);
        
        public List<string> Warnings = new List<string>();

        public List<InputDataFilters_Item> Applicablefilters = new List<InputDataFilters_Item>();
        
        public List<AliasDefinition> AliasDefinitions_BusinessTMP = new List<AliasDefinition>();
        
        public List<AliasDefinition> AliasDefinitions_Categoria = new List<AliasDefinition>();

        public List<SlideToGenerate> SildeToGenerate = new List<SlideToGenerate>();
        
        public List<ItemToExport> ItemsToExportAsImage = new List<ItemToExport>();
        
        public List<string> OutputFilePathLists = new List<string>();









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
            if (input == null) { throw new ArgumentNullException("input"); }

            base.DestinationFolder = input.DestinationFolder;
            base.TmpFolder = input.TmpFolder;
            base.DataSourceFilePath = input.DataSourceFilePath;
            base.DebugFilePath = input.DebugFilePath;
            //
            base.FileBudgetPath = input.FileBudgetPath;
            base.FileForecastPath = input.FileForecastPath;
            base.FileSuperDettagliPath = input.FileSuperDettagliPath;
            base.FileRunRatePath = input.FileRunRatePath;
            //
            ReplaceAllData_FileSuperDettagli = input.ReplaceAllData_FileSuperDettagli;
            PeriodDate = input.PeriodDate;
            Applicablefilters = input.Applicablefilters ?? new List<InputDataFilters_Item>();
        }

        public void SetContextFromInput(ValidateSourceFilesInput input)
        {
            if (input == null) { throw new ArgumentNullException("input"); }

            base.DestinationFolder = input.DestinationFolder;
            base.TmpFolder = input.TmpFolder;
            base.DataSourceFilePath = input.DataSourceFilePath;
            base.DebugFilePath = input.DebugFilePath;
            //
            base.FileBudgetPath = input.FileBudgetPath;
            base.FileForecastPath = input.FileForecastPath;
            base.FileSuperDettagliPath = input.FileSuperDettagliPath;
            base.FileRunRatePath = input.FileRunRatePath;
        }
    }
}