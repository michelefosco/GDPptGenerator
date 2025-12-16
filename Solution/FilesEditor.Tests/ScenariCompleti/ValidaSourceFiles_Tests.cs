using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using FilesEditor.Tests.Constants;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace FilesEditor.Tests
{
    [TestClass]
    public class ValidateSourceFiles_Tests : BaseTest
    {
        string _fileDataSourceName;
        //
        string _fileBudgetName;
        string _fileCN43NtName;
        string _fileForecasttName;
        string _fileRunRatetName;
        string _fileSuperDettaglitName;

        private void SettaDefaults()
        {
            _fileDataSourceName = TestPaths.DataSource_000005_Stay_000002_Go;
            //
            _fileBudgetName = TestPaths.INPUT_BUDGET_FILE;
            _fileCN43NtName = TestPaths.INPUT_CN43N_FILE;
            _fileForecasttName = TestPaths.INPUT_FORECAST_FILE;
            _fileRunRatetName = TestPaths.INPUT_RUNRATE_FILE;
            _fileSuperDettaglitName = TestPaths.INPUT_SUPERDETTAGLI_FILE_00012;
        }

        private ValidateSourceFilesOutput EseguiMetodo()
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
            //var powerPointTemplateFilePath = Path.Combine(BinFolderPath, FileNames.POWERPOINT_TEMPLATE_FILENAME);

            var input = new ValidateSourceFilesInput(
                        dataSourceFilePath: dataSourceFilePath,
                        destinationFolder: destinationFolder,
                        tmpFolder: tmpFolder,
                        debugFilePath: debugFilePath,
                        //
                        fileBudgetPath: fileBudgetPath,
                        fileCN43NPath: fileCN43NPath,
                        fileForecastPath: fleForecastPath,
                        fileRunRatePath: fileRunRatePath,
                        fileSuperDettagliPath: fileSuperDettagliPath
                        );

            return Editor.ValidateSourceFiles(input);
        }


        [TestMethod]
        public void Scenario_OK_01()
        {
            //var output = Editor.ValidateSourceFiles(input);
            SettaDefaults();
            var output = EseguiMetodo();

            // Check generali
            int numeroApplicablefilters = 7;
            InputDataFilters_Tables[] tables = { InputDataFilters_Tables.BUDGET, InputDataFilters_Tables.BUDGET, InputDataFilters_Tables.FORECAST, InputDataFilters_Tables.FORECAST, InputDataFilters_Tables.SUPERDETTAGLI, InputDataFilters_Tables.SUPERDETTAGLI, InputDataFilters_Tables.SUPERDETTAGLI };
            string[] fieldNames = { "Business", "Categoria", "Business", "Categoria", "Last name First name", "Project Description", "Bus Area GDLT" };
            int[] numberOfPossibleValues = { 6, 8, 6, 8, 1, 7, 7 };
            CheckValueFilters(output, numeroApplicablefilters, tables, fieldNames, numberOfPossibleValues);

            // Check specifici
            Assert.AreEqual("CGO & Other div", output.Applicablefilters[0].PossibleValues[0]);
            Assert.AreEqual("AS/REB/SERV", output.Applicablefilters[1].PossibleValues[0]);
            Assert.AreEqual("CGO & Other div", output.Applicablefilters[2].PossibleValues[0]);
            Assert.AreEqual("AS/REB/SERV", output.Applicablefilters[3].PossibleValues[0]);
            Assert.AreEqual("1598", output.Applicablefilters[4].PossibleValues[0]);
            Assert.AreEqual("Proj_1", output.Applicablefilters[5].PossibleValues[0]);
            Assert.AreEqual("Proj_1_name", output.Applicablefilters[6].PossibleValues[0]);
        }

        private static void CheckValueFilters(ValidateSourceFilesOutput output, int numeroApplicablefilters, InputDataFilters_Tables[] tables, string[] fieldNames, int[] numberOfPossibleValues)
        {
            // test base
            Assert.IsNotNull(output);
            Assert.IsNull(output.ManagedException, output.ManagedException?.Message);
            Assert.AreEqual(EsitiFinali.Success, output.Esito);

            // test specifici dell'oggetto di output
            Assert.IsNotNull(output.Applicablefilters);
            Assert.AreEqual(numeroApplicablefilters, output.Applicablefilters.Count);

            foreach (var filter in output.Applicablefilters)
            {
                Assert.IsFalse(string.IsNullOrWhiteSpace(filter.FieldName));
                Assert.IsNotNull(filter.PossibleValues);
                Assert.IsTrue(filter.PossibleValues.Count > 0);
                Assert.AreEqual(0, filter.SelectedValues.Count);
            }


            for (int j = 0; j < numeroApplicablefilters; j++)
            {
                Assert.AreEqual(tables[j], output.Applicablefilters[j].Table);
                Assert.AreEqual(fieldNames[j], output.Applicablefilters[j].FieldName);
                Assert.AreEqual(numberOfPossibleValues[j], output.Applicablefilters[j].PossibleValues.Count);
            }
        }
    }
}
