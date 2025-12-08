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
        // properties base class
        string _PowerPointTemplateFilePath;
        string _DataSourceFolder;
        string _DataSourceFilePath;
        string _DestinationFolder;
        string _TmpFolder;
        string _DebugFilePath;
        // properties specifiche di questo oggetto di input
        string _FileBudgetPath;
        string _FileForecastPath;
        string _FileSuperDettagliPath;
        string _FileRunRatePath;
        string _FileCN43NPath;



        bool _AppendCurrentYear_FileSuperDettagli;
        DateTime _PeriodDate;
        List<InputDataFilters_Item> _Applicablefilters;

        private void SettaDefaults()
        {
            // properties base class
            _PowerPointTemplateFilePath = Path.Combine(BinFolderPath, FileNames.POWERPOINT_TEMPLATE_FILENAME);

            _DataSourceFolder = Path.Combine(TestFileFolderPath, TestPaths.DATASOURCE_FOLDER);

            //_DataSourceFilePath = Path.Combine(_DataSourceFolder, FileNames.DATASOURCE_FILENAME);
            var dataSourceFilePathOriginale = Path.Combine(_DataSourceFolder, FileNames.DATASOURCE_FILENAME);
            _DataSourceFilePath = Path.Combine(_DataSourceFolder, "test_" + FileNames.DATASOURCE_FILENAME);
            File.Copy(dataSourceFilePathOriginale, _DataSourceFilePath, true);

            _DestinationFolder = Path.Combine(TestFileFolderPath, TestPaths.OUTPUT_FOLDER);
            _TmpFolder = Path.Combine(_DestinationFolder, TestPaths.TMP_FOLDER);
            _DebugFilePath = Path.Combine(_DestinationFolder, TestPaths.OUTPUT_DEBUGFILE);

            // properties specifiche di questo oggetto di input
            _FileBudgetPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_BUDGET_FILE);
            _FileForecastPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_FORECAST_FILE);
            _FileSuperDettagliPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_SUPERDETTAGLI_FILE);
            _FileRunRatePath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_RUNRATE_FILE);
            _FileCN43NPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_CN43N_FILE);

            _AppendCurrentYear_FileSuperDettagli = false;
            _PeriodDate = new DateTime(2025, 11, 2);
            _Applicablefilters = new List<InputDataFilters_Item>();
        }

        private UpdataDataSourceAndBuildPresentationOutput EseguiMetodo()
        {
            var input = new UpdataDataSourceAndBuildPresentationInput(
                        dataSourceFilePath: _DataSourceFilePath,
                        destinationFolder: _DestinationFolder,
                        tmpFolder: _TmpFolder,
                        debugFilePath: _DebugFilePath,
                        //
                        fileBudgetPath: _FileBudgetPath,
                        fileForecastPath: _FileForecastPath,
                        fileSuperDettagliPath: _FileSuperDettagliPath,
                        fileRunRatePath: _FileRunRatePath,
                        fileCN43NPath: _FileCN43NPath,
                        //
                        powerPointTemplateFilePath: _PowerPointTemplateFilePath,
                        appendCurrentYear_FileSuperDettagli: _AppendCurrentYear_FileSuperDettagli,
                        periodDate: _PeriodDate,
                        applicablefilters: _Applicablefilters
                        );
            return Editor.UpdataDataSourceAndBuildPresentation(input);
        }

        private void CheckResults(UpdataDataSourceAndBuildPresentationOutput output, string dataSourceFilePath, string tmpFolder, int numeroRigheBudget, int numeroRigheForecast, int numeroRigheSuperdettagli, int numeroRigheRunRate, int numeroRigheCN43N, int numeroFilesFotoInTmpFolder, int numeroWarnings, int numeroPresentazioniGenerate, int numeroRigheAnnoCorrente)
        {
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
        public void Scenario_OK_001()
        {
            SettaDefaults();

            // Personalizzazione parametri

            var output = EseguiMetodo();
            CheckResults(
                output: output,
                dataSourceFilePath: _DataSourceFilePath,
                tmpFolder: _TmpFolder,
                // solo le righe effettiva, senza considerare le intestazine e le righe in alto
                numeroRigheBudget: 34,
                numeroRigheForecast: 34,
                numeroRigheSuperdettagli: 1000,
                numeroRigheRunRate: 1,
                numeroRigheCN43N: 12,
                //
                numeroFilesFotoInTmpFolder: 14,
                numeroWarnings: 0,
                numeroPresentazioniGenerate: 2,
                //
                numeroRigheAnnoCorrente: 1000
                );
        }

        [TestMethod]
        public void Scenario_OK_002()
        {
            SettaDefaults();

            // Personalizzazione parametri
            _AppendCurrentYear_FileSuperDettagli = true;

            var output = EseguiMetodo();
            CheckResults(
                output: output,
                dataSourceFilePath: _DataSourceFilePath,
                tmpFolder: _TmpFolder,
                // solo le righe effettiva, senza considerare le intestazine e le righe in alto
                numeroRigheBudget: 34,
                numeroRigheForecast: 34,
                numeroRigheSuperdettagli: 1000 + 5, // 4 righe con anni diversi dal 2025
                numeroRigheRunRate: 1,
                numeroRigheCN43N: 12,
                //
                numeroFilesFotoInTmpFolder: 14,
                numeroWarnings: 0,
                numeroPresentazioniGenerate: 2,
                //
                numeroRigheAnnoCorrente: 1000
                );
        }

        [TestMethod]
        public void Scenario_OK_003()
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
                dataSourceFilePath: _DataSourceFilePath,
                tmpFolder: _TmpFolder,
                // solo le righe effettiva, senza considerare le intestazine e le righe in alto
                numeroRigheBudget: 1,
                numeroRigheForecast: 34,
                numeroRigheSuperdettagli: 1000,
                numeroRigheRunRate: 1,
                numeroRigheCN43N: 12,
                //
                numeroFilesFotoInTmpFolder: 14,
                numeroWarnings: 0,
                numeroPresentazioniGenerate: 2,
                //
                numeroRigheAnnoCorrente: 1000
                );
        }

        [TestMethod]
        public void Scenario_OK_004()
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
                dataSourceFilePath: _DataSourceFilePath,
                tmpFolder: _TmpFolder,
                // solo le righe effettiva, senza considerare le intestazine e le righe in alto
                numeroRigheBudget: 1,
                numeroRigheForecast: 7,
                numeroRigheSuperdettagli: 1000,
                numeroRigheRunRate: 1,
                numeroRigheCN43N: 12,
                //
                numeroFilesFotoInTmpFolder: 14,
                numeroWarnings: 0,
                numeroPresentazioniGenerate: 2,
                //
                numeroRigheAnnoCorrente: 1000
                );
        }

        [TestMethod]
        public void Scenario_OK_005()
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
                dataSourceFilePath: _DataSourceFilePath,
                tmpFolder: _TmpFolder,
                // solo le righe effettiva, senza considerare le intestazine e le righe in alto
                numeroRigheBudget: 1,
                numeroRigheForecast: 4,
                numeroRigheSuperdettagli: 1000,
                numeroRigheRunRate: 1,
                numeroRigheCN43N: 12,
                //
                numeroFilesFotoInTmpFolder: 14,
                numeroWarnings: 0,
                numeroPresentazioniGenerate: 2,
                //
                numeroRigheAnnoCorrente: 1000
                );
        }


        [TestMethod]
        public void Scenario_OK_006()
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
                dataSourceFilePath: _DataSourceFilePath,
                tmpFolder: _TmpFolder,
                // solo le righe effettiva, senza considerare le intestazine e le righe in alto
                numeroRigheBudget: 1,
                numeroRigheForecast: 1,
                numeroRigheSuperdettagli: 1000,
                numeroRigheRunRate: 1,
                numeroRigheCN43N: 12,
                //
                numeroFilesFotoInTmpFolder: 14,
                numeroWarnings: 0,
                numeroPresentazioniGenerate: 2,
                //
                numeroRigheAnnoCorrente: 1000
                );
        }
    }
}
