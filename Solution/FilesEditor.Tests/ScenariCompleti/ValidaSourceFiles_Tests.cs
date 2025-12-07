using FilesEditor.Entities;
using FilesEditor.Entities.MethodsArgs;
using FilesEditor.Enums;
using FilesEditor.Tests.Constants;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;

namespace FilesEditor.Tests
{
    [TestClass]
    public class ValidateSourceFiles_Tests : BaseTest
    {
        [TestMethod]
        public void Scenario_OK_01()
        {
            // properties base class
            string dataSourceFolder = Path.Combine(TestFileFolderPath, TestPaths.DATASOURCE_FOLDER);
            string dataSourceFilePath = Path.Combine(dataSourceFolder, FilesEditor.Constants.FileNames.DATASOURCE_FILENAME);
            string destinationFolder = Path.Combine(TestFileFolderPath, TestPaths.OUTPUT_FOLDER);
            string tmpFolder = Path.Combine(destinationFolder, TestPaths.TMP_FOLDER);
            string debugFilePath = Path.Combine(destinationFolder, TestPaths.OUTPUT_DEBUGFILE);

            // properties specifiche di questo oggetto di input
            string fileBudgetPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_BUDGET_FILE);
            string fileForecastPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_FORECAST_FILE);
            string fileSuperDettagliPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_SUPERDETTAGLI_FILE);
            string fileRunRatePath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_RUNRATE_FILE);
            string fileCN43NPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_CN43N_FILE);

            var input = new ValidateSourceFilesInput(
                    dataSourceFilePath: dataSourceFilePath,
                    destinationFolder: destinationFolder,
                    tmpFolder: tmpFolder,
                    debugFilePath: debugFilePath,
                    //
                    fileBudgetPath: fileBudgetPath,
                    fileForecastPath: fileForecastPath,
                    fileSuperDettagliPath: fileSuperDettagliPath,
                    fileRunRatePath: fileRunRatePath,
                    fileCN43NPath: fileCN43NPath
                    );

            var output = Editor.ValidateSourceFiles(input);


            // Check generali
            int numeroApplicablefilters = 7;
            InputDataFilters_Tables[] tables = { InputDataFilters_Tables.BUDGET, InputDataFilters_Tables.BUDGET, InputDataFilters_Tables.FORECAST, InputDataFilters_Tables.FORECAST, InputDataFilters_Tables.SUPERDETTAGLI, InputDataFilters_Tables.SUPERDETTAGLI, InputDataFilters_Tables.SUPERDETTAGLI };
            string[] fieldNames = { "Business", "Categoria", "Business", "Categoria", "Last name First name", "Project Description", "Bus Area GDLT" };
            int[] numberOfPossibleValues = { 6, 8, 6, 8, 1, 10, 11 };
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
