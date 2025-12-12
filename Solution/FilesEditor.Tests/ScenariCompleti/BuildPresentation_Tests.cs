using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using FilesEditor.Tests.Constants;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;

namespace FilesEditor.Tests
{
    [TestClass]
    public class BuildPresentation_Tests : BaseTest
    {
        string _fileDataSourceName;
        //
        string _fileBudgetName;
        string _fileCN43NtName;
        string _fileForecasttName;
        string _fileRunRatetName;
        string _fileSuperDettaglitName;
        //
        bool _AppendCurrentYear_FileSuperDettagli;
        DateTime _PeriodDate;
        List<InputDataFilters_Item> _Applicablefilters;

        private void SettaDefaults()
        {
            _fileDataSourceName = TestPaths.DATASOURCE_005Stay_002Go_FILENAME;
            //
            _fileBudgetName = TestPaths.INPUT_BUDGET_FILE;
            _fileCN43NtName = TestPaths.INPUT_CN43N_FILE;
            _fileForecasttName = TestPaths.INPUT_FORECAST_FILE;
            _fileRunRatetName = TestPaths.INPUT_RUNRATE_FILE;
            _fileSuperDettaglitName = TestPaths.INPUT_SUPERDETTAGLI_FILE_012;
            //
            _AppendCurrentYear_FileSuperDettagli = true;
            _PeriodDate = new DateTime(2025, 11, 2);
            _Applicablefilters = new List<InputDataFilters_Item>();
        }

        private UpdataDataSourceAndBuildPresentationOutput EseguiMetodo()
        {
            var destinationFolder = Path.Combine(TestFileFolderPath, TestPaths.OUTPUT_FOLDER);
            var tmpFolder = Path.Combine(destinationFolder, TestPaths.TMP_FOLDER);
            var debugFilePath = Path.Combine(destinationFolder, TestPaths.OUTPUT_DEBUGFILE);
            //
            var dataSourceFolder = Path.Combine(TestFileFolderPath, TestPaths.DATASOURCE_FOLDER);
            var dataSourceFilePathOriginale = Path.Combine(dataSourceFolder, _fileDataSourceName);
            var dataSourceFilePath = Path.Combine(dataSourceFolder, TestPaths.DATASOURCE_TEST_FILENAME);
            File.Copy(dataSourceFilePathOriginale, dataSourceFilePath, true);
            //
            var inputFilesFolder = Path.Combine(TestFileFolderPath, TestPaths.INPUTFILES_FOLDER);
            //
            var fileBudgetPath = Path.Combine(inputFilesFolder, _fileBudgetName);
            var fileCN43NPath = Path.Combine(inputFilesFolder, _fileCN43NtName);
            var fleForecastPath = Path.Combine(inputFilesFolder, _fileForecasttName);
            var fileRunRatePath = Path.Combine(inputFilesFolder, _fileRunRatetName);
            var fileSuperDettagliPath = Path.Combine(inputFilesFolder, _fileSuperDettaglitName);
            //
            var powerPointTemplateFilePath = Path.Combine(BinFolderPath, FileNames.POWERPOINT_TEMPLATE_FILENAME);

            var input = new UpdataDataSourceAndBuildPresentationInput(
                        dataSourceFilePath: dataSourceFilePath,
                        destinationFolder: destinationFolder,
                        tmpFolder: tmpFolder,
                        debugFilePath: debugFilePath,
                        //
                        fileBudgetPath: fileBudgetPath,
                        fileCN43NPath: fileCN43NPath,
                        fileForecastPath: fleForecastPath,
                        fileRunRatePath: fileRunRatePath,
                        fileSuperDettagliPath: fileSuperDettagliPath,
                        //
                        powerPointTemplateFilePath: powerPointTemplateFilePath,
                        //
                        appendCurrentYear_FileSuperDettagli: _AppendCurrentYear_FileSuperDettagli,
                        periodDate: _PeriodDate,
                        applicablefilters: _Applicablefilters
                        );
            return Editor.UpdataDataSourceAndBuildPresentation(input);
        }

        private void CheckResults(UpdataDataSourceAndBuildPresentationOutput output, int numeroRigheBudget, int numeroRigheForecast, int numeroRigheSuperdettagli, int numeroRigheRunRate, int numeroRigheCN43N, int numeroFilesFotoInTmpFolder, int numeroWarnings, int numeroPresentazioniGenerate, int numeroRigheAnnoCorrente)
        {
            var destinationFolder = Path.Combine(TestFileFolderPath, TestPaths.OUTPUT_FOLDER);
            var tmpFolder = Path.Combine(destinationFolder, TestPaths.TMP_FOLDER);
            var dataSourceFolder = Path.Combine(TestFileFolderPath, TestPaths.DATASOURCE_FOLDER);
            var dataSourceFilePath = Path.Combine(dataSourceFolder, TestPaths.DATASOURCE_TEST_FILENAME);

            // test base
            Assert.IsNotNull(output);
            Assert.IsNull(output.ManagedException);
            Assert.AreEqual(EsitiFinali.Success, output.Esito);


            // numero file di output generati
            Assert.AreEqual(numeroPresentazioniGenerate, output.OutputFilePathLists.Count);

            // numero di warnings sollevati
            Assert.AreEqual(numeroWarnings, output.Warnings.Count);
            //Assert.IsTrue(output.Warnings.Any(_ => _.Contains("The input file 'Super dettagli' contains at least one year that is different from the year selected as the period date")));

            // con la cancellazione del tmp folder non si può più verificare
            // numero di immagini generate su file system
            //var filesInTmpFolder = Directory.GetFiles(tmpFolder);
            //Assert.AreEqual(numeroFilesFotoInTmpFolder, filesInTmpFolder.Length);

            var ePPlusHelper = new EPPlusExtensions.EPPlusHelper();
            ePPlusHelper.Open(dataSourceFilePath);



            const int offSetIntestazioneBudget = 2;
            Assert.AreEqual(offSetIntestazioneBudget + numeroRigheBudget + 1,   // expected
                            ePPlusHelper.GetFirstEmptyRow(WorksheetNames.DATASOURCE_BUDGET_DATA, offSetIntestazioneBudget, 1));
            //
            const int offSetIntestazioneForecast = 2;
            Assert.AreEqual(offSetIntestazioneForecast + numeroRigheForecast + 1,   // expected
                            ePPlusHelper.GetFirstEmptyRow(WorksheetNames.DATASOURCE_FORECAST_DATA, offSetIntestazioneForecast, 1));
            //
            const int offSetIntestazioneRunRate = 1;
            Assert.AreEqual(offSetIntestazioneRunRate + numeroRigheRunRate + 1,   // expected
                            ePPlusHelper.GetFirstEmptyRow(WorksheetNames.DATASOURCE_RUN_RATE_DATA, offSetIntestazioneRunRate, 1));
            //
            const int offSetIntestazioneSuperdettagli = 2;
            var numeroRigheInSuperdettagliInput = ePPlusHelper.GetFirstEmptyRow(WorksheetNames.DATASOURCE_SUPERDETTAGLI_DATA, offSetIntestazioneSuperdettagli, 1);
            Assert.AreEqual(offSetIntestazioneSuperdettagli + numeroRigheSuperdettagli + 1,   // expected
                            numeroRigheInSuperdettagliInput);
            //
            const int offSetIntestazioneRigheCN43N = 1;
            Assert.AreEqual(offSetIntestazioneRigheCN43N + numeroRigheCN43N + 1,   // expected
                            ePPlusHelper.GetFirstEmptyRow(WorksheetNames.DATASOURCE_CN43N_DATA, offSetIntestazioneRigheCN43N, 1));

            // verifico il numero di righe con l'anno corrispondente al periodo selezionato
            var ws = ePPlusHelper.ExcelPackage.Workbook.Worksheets[WorksheetNames.DATASOURCE_SUPERDETTAGLI_DATA];
            Assert.AreEqual(numeroRigheAnnoCorrente, // Expected
                            ws.Cells[2, // Prima riga,
                                        2,  // Colonna Anno
                                        ws.Dimension.End.Row, //Ultima riga
                                        2  // Colonna Anno
                                        ].Select(c => c.Text).Count(_ => string.Equals(_, _PeriodDate.Year.ToString(), StringComparison.Ordinal)) // Actual
            );
        }





        [TestMethod]
        public void Scenario_OK_001_Importo_PIU_RigheDiQuelleCancellate()
        {
            SettaDefaults();

            // Personalizzazione parametri

            // 2 righe vanno e 12 se ne aggiungono

            var output = EseguiMetodo();
            CheckResults(
                output: output,
                // solo le righe effettiva, senza considerare le intestazine e le righe in alto
                numeroRigheBudget: 34,
                numeroRigheForecast: 34,
                numeroRigheSuperdettagli: 5 + 12, // 5 stay from different years + 12 added by appending current year
                numeroRigheRunRate: 1,
                numeroRigheCN43N: 12,
                //
                numeroFilesFotoInTmpFolder: 14,
                numeroWarnings: 0,
                numeroPresentazioniGenerate: 2,
                //
                numeroRigheAnnoCorrente: 12
                );
        }

        [TestMethod]
        public void Scenario_OK_001_Importo_MENO_RigheDiQuelleCancellate()
        {
            SettaDefaults();

            // Personalizzazione parametri
            _fileDataSourceName = TestPaths.DATASOURCE_010Stay_020Go_FILENAME;

            // 20 righe vanno e 12 se ne aggiungono
            var output = EseguiMetodo();
            CheckResults(
                output: output,
                // solo le righe effettiva, senza considerare le intestazine e le righe in alto
                numeroRigheBudget: 34,
                numeroRigheForecast: 34,
                numeroRigheSuperdettagli: 10 + 12, // 10 stay from different years + 12 added by appending current year
                numeroRigheRunRate: 1,
                numeroRigheCN43N: 12,
                //
                numeroFilesFotoInTmpFolder: 14,
                numeroWarnings: 0,
                numeroPresentazioniGenerate: 2,
                //
                numeroRigheAnnoCorrente: 12
                );
        }

        [TestMethod]
        public void Scenario_OK_001_Append()
        {
            SettaDefaults();

            // Personalizzazione parametri
            _AppendCurrentYear_FileSuperDettagli = false;

            var output = EseguiMetodo();
            CheckResults(
                output: output,
                // solo le righe effettiva, senza considerare le intestazine e le righe in alto
                numeroRigheBudget: 34,
                numeroRigheForecast: 34,
                numeroRigheSuperdettagli: 0 + 12, // 0 stay from different years + 12 added by appending current year
                numeroRigheRunRate: 1,
                numeroRigheCN43N: 12,
                //
                numeroFilesFotoInTmpFolder: 14,
                numeroWarnings: 0,
                numeroPresentazioniGenerate: 2,
                //
                numeroRigheAnnoCorrente: 12
                );
        }

        [TestMethod]
        public void Scenario_OK_WithFileds_001()
        {
            SettaDefaults();

            // Personalizzazione parametri
            _Applicablefilters = new List<InputDataFilters_Item>
            {
                new InputDataFilters_Item {
                    Table = InputDataFilters_Tables.BUDGET,
                    FieldName = Values.HEADER_BUSINESS,
                    SelectedValues = new List<string>{ "TOB (MK, PK, MO)" }
                },
                new InputDataFilters_Item {
                    Table = InputDataFilters_Tables.BUDGET,
                    FieldName = Values.HEADER_CATEGORIA,
                    SelectedValues = new List<string>{ "Match caso Macchina" }
                },

                //new InputDataFilters_Item {
                //    Table = InputDataFilters_Tables.FORECAST,
                //    FieldName = Values.HEADER_BUSINESS,
                //    SelectedValues = new List<string>{ "ESS" }
                //},
                //new InputDataFilters_Item {
                //    Table = InputDataFilters_Tables.FORECAST,
                //    FieldName = Values.HEADER_CATEGORIA,
                //    SelectedValues = new List<string>{ "FEASIBILITY" }
                //}
            };

            var output = EseguiMetodo();
            CheckResults(
                output: output,
                // solo le righe effettiva, senza considerare le intestazine e le righe in alto
                numeroRigheBudget: 1,
                numeroRigheForecast: 34,
                numeroRigheSuperdettagli: 5 + 12, // 5 stay from different years + 12 added by appending current year
                numeroRigheRunRate: 1,
                numeroRigheCN43N: 12,
                //
                numeroFilesFotoInTmpFolder: 14,
                numeroWarnings: 0,
                numeroPresentazioniGenerate: 2,
                //
                numeroRigheAnnoCorrente: 12
                );
        }

        [TestMethod]
        public void Scenario_OK_WithFileds_002()
        {
            SettaDefaults();

            // Personalizzazione parametri
            _Applicablefilters = new List<InputDataFilters_Item>
            {
                new InputDataFilters_Item {
                    Table = InputDataFilters_Tables.BUDGET,
                    FieldName = Values.HEADER_BUSINESS,
                    SelectedValues = new List<string>{ "TOB (MK, PK, MO)" }
                },
                new InputDataFilters_Item {
                    Table = InputDataFilters_Tables.BUDGET,
                    FieldName = Values.HEADER_CATEGORIA,
                    SelectedValues = new List<string>{ "Match caso Macchina" }
                },
                new InputDataFilters_Item {
                    Table = InputDataFilters_Tables.FORECAST,
                    FieldName = Values.HEADER_BUSINESS,
                    SelectedValues = new List<string>{ "ESS" }
                },
                //new InputDataFilters_Item {
                //    Table = InputDataFilters_Tables.FORECAST,
                //    FieldName = Values.HEADER_CATEGORIA,
                //    SelectedValues = new List<string>{ "FEASIBILITY" }
                //}
            };

            var output = EseguiMetodo();
            CheckResults(
                output: output,
                // solo le righe effettiva, senza considerare le intestazine e le righe in alto
                numeroRigheBudget: 1,
                numeroRigheForecast: 7,
                numeroRigheSuperdettagli: 5 + 12, // 5 stay from different years + 12 added by appending current year
                numeroRigheRunRate: 1,
                numeroRigheCN43N: 12,
                //
                numeroFilesFotoInTmpFolder: 14,
                numeroWarnings: 0,
                numeroPresentazioniGenerate: 2,
                //
                numeroRigheAnnoCorrente: 12
                );
        }

        [TestMethod]
        public void Scenario_OK_WithFileds_003()
        {
            SettaDefaults();

            // Personalizzazione parametri
            _Applicablefilters = new List<InputDataFilters_Item>
            {
                new InputDataFilters_Item {
                    Table = InputDataFilters_Tables.BUDGET,
                    FieldName = Values.HEADER_BUSINESS,
                    SelectedValues = new List<string>{ "TOB (MK, PK, MO)" }
                },
                new InputDataFilters_Item {
                    Table = InputDataFilters_Tables.BUDGET,
                    FieldName = Values.HEADER_CATEGORIA,
                    SelectedValues = new List<string>{ "Match caso Macchina" }
                },
                //new InputDataFilters_Item {
                //    Table = InputDataFilters_Tables.FORECAST,
                //    FieldName = Values.HEADER_BUSINESS,
                //    SelectedValues = new List<string>{ "ESS" }
                //},
                new InputDataFilters_Item {
                    Table = InputDataFilters_Tables.FORECAST,
                    FieldName = Values.HEADER_CATEGORIA,
                    SelectedValues = new List<string>{ "FEASIBILITY" }
                }
            };

            var output = EseguiMetodo();
            CheckResults(
                output: output,
                // solo le righe effettiva, senza considerare le intestazine e le righe in alto
                numeroRigheBudget: 1,
                numeroRigheForecast: 4,
                numeroRigheSuperdettagli: 5 + 12, // 5 stay from different years + 12 added by appending current year
                numeroRigheRunRate: 1,
                numeroRigheCN43N: 12,
                //
                numeroFilesFotoInTmpFolder: 14,
                numeroWarnings: 0,
                numeroPresentazioniGenerate: 2,
                //
                numeroRigheAnnoCorrente: 12
                );
        }

        [TestMethod]
        public void Scenario_OK_WithFileds_004()
        {
            SettaDefaults();

            // Personalizzazione parametri
            _Applicablefilters = new List<InputDataFilters_Item>
            {
                new InputDataFilters_Item {
                    Table = InputDataFilters_Tables.BUDGET,
                    FieldName = Values.HEADER_BUSINESS,
                    SelectedValues = new List<string>{ "TOB (MK, PK, MO)" }
                },
                new InputDataFilters_Item {
                    Table = InputDataFilters_Tables.BUDGET,
                    FieldName = Values.HEADER_CATEGORIA,
                    SelectedValues = new List<string>{ "Match caso Macchina" }
                },
                new InputDataFilters_Item {
                    Table = InputDataFilters_Tables.FORECAST,
                    FieldName = Values.HEADER_BUSINESS,
                    SelectedValues = new List<string>{ "ESS" }
                },
                new InputDataFilters_Item {
                    Table = InputDataFilters_Tables.FORECAST,
                    FieldName = Values.HEADER_CATEGORIA,
                    SelectedValues = new List<string>{ "FEASIBILITY" }
                }
            };

            var output = EseguiMetodo();
            CheckResults(
                output: output,
                // solo le righe effettiva, senza considerare le intestazine e le righe in alto
                numeroRigheBudget: 1,
                numeroRigheForecast: 1,
                numeroRigheSuperdettagli: 5 + 12, // 5 stay from different years + 12 added by appending current year
                numeroRigheRunRate: 1,
                numeroRigheCN43N: 12,
                //
                numeroFilesFotoInTmpFolder: 14,
                numeroWarnings: 0,
                numeroPresentazioniGenerate: 2,
                //
                numeroRigheAnnoCorrente: 12
                );
        }

        [TestMethod]
        public void Scenario_OK_Big_Files_001()
        {
            SettaDefaults();

            // Personalizzazione parametri
            //_fileDataSourceName = TestPaths.DATASOURCE_k230Stay_k40Go_FILENAME;
            _fileSuperDettaglitName = TestPaths.INPUT_SUPERDETTAGLI_FILE_F40;


            var output = EseguiMetodo();
            CheckResults(
                output: output,
                // solo le righe effettiva, senza considerare le intestazine e le righe in alto
                numeroRigheBudget: 34,
                numeroRigheForecast: 34,
                numeroRigheSuperdettagli: 5 + 40000, // 5 stay from different years + 12 added by appending current year
                numeroRigheRunRate: 1,
                numeroRigheCN43N: 12,
                //
                numeroFilesFotoInTmpFolder: 14,
                numeroWarnings: 0,
                numeroPresentazioniGenerate: 2,
                //
                numeroRigheAnnoCorrente: 40000
                );
        }
    }
}
