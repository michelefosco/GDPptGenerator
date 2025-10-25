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
        private EPPlusHelper _ePPlusHelperDataSource;

        // Base class properties
        // DestinationFolder
        // TmpFolder
        // DataSourceFilePath
        // DebugFilePath
        // FileBudgetPath
        // FileForecastPath
        // FileSuperDettagliPath
        // FileRunRatePath

        // Input specifico di uno o più metodi
        public bool ReplaceAllData_FileSuperDettagli { get; private set; }
        public string PowerPointTemplateFilePath { get; private set; }

        #region Period
        public DateTime PeriodDate { get; private set; }
        public int PeriodYear { get { return PeriodDate.Year; } }
        public int PeriodMont { get { return PeriodDate.Month; } }
        public int PeriodQuarter { get { return (int)((PeriodDate.Month + 2) / 3); } }
        #endregion

        public Configurazione Configurazione { get; private set; }
        public EsitiFinali Esito { get; private set; }




        public EPPlusHelper EpplusHelperDataSource
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

        public DebugInfoLogger DebugInfoLogger { get; private set; }

        public List<string> Warnings { get; private set; }

        public List<InputDataFilters_Item> ApplicableFilters { get; private set; }

        public List<AliasDefinition> AliasDefinitions_BusinessTMP { get; private set; }

        public List<AliasDefinition> AliasDefinitions_Categoria { get; private set; }

        public List<SlideToGenerate> SildeToGenerate { get; private set; }

        public List<ItemToExport> ItemsToExportAsImage { get; private set; }

        public List<string> OutputFilePathLists { get; private set; }


        public StepContext(Configurazione configurazione)
        {
            Configurazione = configurazione;
            //
            DebugInfoLogger = new DebugInfoLogger(null);
            Warnings = new List<string>();
            ApplicableFilters = new List<InputDataFilters_Item>();
            AliasDefinitions_BusinessTMP = new List<AliasDefinition>();
            AliasDefinitions_Categoria = new List<AliasDefinition>();
            SildeToGenerate = new List<SlideToGenerate>();
            ItemsToExportAsImage = new List<ItemToExport>();
            OutputFilePathLists = new List<string>();
        }


        public void SettaEsitoFinale(EsitiFinali esito)
        {
            Esito = esito;
        }
        public void SetDebugInfoLogger(DebugInfoLogger debugInfoLogger)
        {
            DebugInfoLogger = debugInfoLogger;
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
            PowerPointTemplateFilePath= input.PowerPointTemplateFilePath;
            ReplaceAllData_FileSuperDettagli = input.ReplaceAllData_FileSuperDettagli;
            PeriodDate = input.PeriodDate;
            ApplicableFilters = input.ApplicableFilters ?? new List<InputDataFilters_Item>();
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