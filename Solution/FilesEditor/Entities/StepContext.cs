using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FilesEditor.Entities
{
    public class StepContext : UserInterfaceInputBase
    {
        private EPPlusHelper _dataSourceEPPlusHelper;
        private EPPlusHelper _budgetFileEPPlusHelper;
        private EPPlusHelper _forecastFileEPPlusHelper;
        private EPPlusHelper _runRateFileEPPlusHelper;
        private EPPlusHelper _superdettagliFileEPPlusHelper;
        private EPPlusHelper _cN43NFileEPPlusHelper;


        // Base class properties
        // DestinationFolder
        // TmpFolder
        // DataSourceFilePath
        // DebugFilePath
        // FileBudgetPath
        // FileForecastPath
        // FileSuperDettagliPath
        // FileRunRatePath
        // FileCN43NPath

        // Input specifico di uno o più metodi
        public bool AppendCurrentYear_FileSuperDettagli { get; private set; }
        public string PowerPointTemplateFilePath { get; private set; }

        #region Period
        public DateTime PeriodDate { get; private set; }
        public int PeriodYear { get { return PeriodDate.Year; } }
        public int PeriodMont { get { return PeriodDate.Month; } }
        public string PeriodQuarter { get { return $"Q{(int)((PeriodDate.Month + 2) / 3)}"; } }
        #endregion

        public Configurazione Configurazione { get; private set; }
        public EsitiFinali Esito { get; private set; }

        public EPPlusHelper DataSourceEPPlusHelper
        {
            get
            {
                if (_dataSourceEPPlusHelper == null)
                {
                    if (string.IsNullOrEmpty(DataSourceFilePath))
                    { throw new Exception("Inizializzare 'DataSourceFilePath' prima di usare 'DataSourceEPPlusHelper'"); }

                    _dataSourceEPPlusHelper = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(DataSourceFilePath, FileTypes.DataSource);
                }
                return _dataSourceEPPlusHelper;
            }
        }
        public EPPlusHelper BudgetFileEPPlusHelper
        {
            get
            {
                if (_budgetFileEPPlusHelper == null)
                {
                    if (string.IsNullOrEmpty(FileBudgetPath))
                    { throw new Exception("Inizializzare 'FileBudgetPath' prima di usare 'BudgetFileEPPlusHelper'"); }

                    _budgetFileEPPlusHelper = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(FileBudgetPath, FileTypes.Budget);
                    
                }
                return _budgetFileEPPlusHelper;
            }
        }
        public EPPlusHelper ForecastFileEPPlusHelper
        {
            get
            {
                if (_forecastFileEPPlusHelper == null)
                {
                    if (string.IsNullOrEmpty(FileForecastPath))
                    { throw new Exception("Inizializzare 'FileForecastPath' prima di usare 'ForecastFileEPPlusHelper'"); }

                    _forecastFileEPPlusHelper = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(FileForecastPath, FileTypes.Forecast);
                }
                return _forecastFileEPPlusHelper;
            }
        }
        public EPPlusHelper RunRateFileEPPlusHelper
        {
            get
            {
                if (_runRateFileEPPlusHelper == null)
                {
                    if (string.IsNullOrEmpty(FileRunRatePath))
                    { throw new Exception("Inizializzare 'FileRunRatePath' prima di usare 'RunRateFileEPPlusHelper'"); }

                    _runRateFileEPPlusHelper = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(FileRunRatePath, FileTypes.RunRate);
                }
                return _runRateFileEPPlusHelper;
            }
        }
        public EPPlusHelper SuperdettagliFileEPPlusHelper
        {
            get
            {
                if (_superdettagliFileEPPlusHelper == null)
                {
                    if (string.IsNullOrEmpty(FileSuperDettagliPath))
                    { throw new Exception("Inizializzare 'FileSuperDettagliPath' prima di usare 'SuperdettagliFileEPPlusHelper'"); }

                    _superdettagliFileEPPlusHelper = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(FileSuperDettagliPath, FileTypes.SuperDettagli);
                }
                return _superdettagliFileEPPlusHelper;
            }
        }

        public EPPlusHelper CN43NFileEPPlusHelper
        {
            get
            {
                if (_cN43NFileEPPlusHelper == null)
                {
                    if (string.IsNullOrEmpty(FileCN43NPath))
                    { throw new Exception("Inizializzare 'FileCN43NPath' prima di usare 'CN43NFileEPPlusHelper'"); }

                    _cN43NFileEPPlusHelper = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(FileCN43NPath, FileTypes.CN43N);
                }
                return _cN43NFileEPPlusHelper;
            }
        }

        

        public DebugInfoLogger DebugInfoLogger { get; private set; }

        public List<string> Warnings { get; private set; }

        public List<InputDataFilters_Item> ApplicableFilters { get; private set; }

        public List<AliasDefinition> AliasDefinitions_Business { get; private set; }

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
            AliasDefinitions_Business = new List<AliasDefinition>();
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
            base.FileCN43NPath = input.FileCN43NPath;
            
            //
            PowerPointTemplateFilePath = input.PowerPointTemplateFilePath;
            AppendCurrentYear_FileSuperDettagli = input.AppendCurrentYear_FileSuperDettagli;
            PeriodDate = input.PeriodDate;
            ApplicableFilters = new List<InputDataFilters_Item>(input.ApplicableFilters)
                             ?? new List<InputDataFilters_Item>();
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
            base.FileCN43NPath = input.FileCN43NPath;
        }
        internal void AddWarning(string warningMessage)
        {
            Warnings.Add(warningMessage);
            DebugInfoLogger?.LogWarning(warningMessage);
        }
        internal string ApplicaAliasToValue(string header, string value)
        {
            if (string.IsNullOrEmpty(header))
            { return value; }

            #region Controllo se il valore appartiene ad uno di quelli interessati da Aliases
            List<AliasDefinition> aliasesToCheck = null;
            if (header.Equals(Values.HEADER_BUSINESS, StringComparison.InvariantCultureIgnoreCase))
            {
                aliasesToCheck = AliasDefinitions_Business;
            }
            else if (header.Equals(Values.HEADER_CATEGORIA, StringComparison.InvariantCultureIgnoreCase))
            {
                aliasesToCheck = AliasDefinitions_Categoria;
            }
            else
            {
                // non è necessario applicare gli alias
                return value;
            }
            #endregion

            // non ci sono aliases da verificare
            if (aliasesToCheck == null || aliasesToCheck.Count == 0)
            { return value; }

            #region Cerco un alias che corrisponda
            // Controllo prima gli aliases fissi (senza regular expressions)
            foreach (AliasDefinition alias in aliasesToCheck.Where(_ => !_.IsRegularExpression))
            {
                if (alias.RawValue.Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase))
                { return alias.NewValue; }
            }

            // Controllo successivamente gli aliases con regular expressions
            foreach (AliasDefinition alias in aliasesToCheck.Where(_ => _.IsRegularExpression))
            {
                if (ValuesHelper.StringMatch(value, alias.RawValue))
                { return alias.NewValue; }
            }
            #endregion

            // Nessun alias applicato
            return value;
        }
    }
}