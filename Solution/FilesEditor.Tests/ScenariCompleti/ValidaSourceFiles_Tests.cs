using FilesEditor.Entities;
using FilesEditor.Enums;
using FilesEditor.Tests.Constants;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace FilesEditor.Tests
{
    [TestClass]
    public class ValidateSourceFiles_Tests : BaseTest
    {
        [TestMethod]
        public void Scenario_OK_01()
        {
            // properties base class
            string dataSourceFolder = Path.Combine(TestFileFolderPath, TestPaths.SOURCEFILE_FOLDER);
            string dataSourceFilePath = Path.Combine(dataSourceFolder, FilesEditor.Constants.FileNames.DATASOURCE_FILENAME);
            string destinationFolder = Path.Combine(TestFileFolderPath, TestPaths.OUTPUT_FOLDER);
            string tmpFolder = Path.Combine(destinationFolder, TestPaths.TMP_FOLDER);
            string debugFilePath = Path.Combine(destinationFolder, TestPaths.OUTPUT_DEBUGFILE);

            // properties specifiche di questo oggetto di input
            string fileBudgetPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_BUDGET_FILE);
            string fileForecastPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_FORECAST_FILE);
            string fileSuperDettagliPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_SUPERDETTAGLI_FILE);
            string fileRunRatePat = Path.Combine(TestFileFolderPath, TestPaths.INPUT_RUNRATE_FILE);


            var input = new ValidateSourceFilesInput(
                    dataSourceFilePath: dataSourceFilePath,
                    destinationFolder: destinationFolder,
                    tmpFolder: tmpFolder,
                    debugFilePath: debugFilePath,
                    //
                    fileBudgetPath: fileBudgetPath,
                    fileForecastPath: fileForecastPath,
                    fileSuperDettagliPath: fileSuperDettagliPath,
                    fileRunRatePath: fileRunRatePat
                    );

            var output = Editor.ValidateSourceFiles(input);

            // test base
            Assert.IsNotNull(output);
            Assert.IsNull(output.ManagedException, output.ManagedException?.Message);
            Assert.AreEqual(EsitiFinali.Success, output.Esito);

            // test specifici dell'oggetto di output
            Assert.IsNotNull(output.Applicablefilters);
            Assert.AreEqual(7, output.Applicablefilters.Count);

            foreach (var filter in output.Applicablefilters)
            {
                Assert.IsFalse(string.IsNullOrWhiteSpace(filter.FieldName));
                Assert.IsNotNull(filter.Values);
                Assert.IsTrue(filter.Values.Count > 0);
                Assert.AreEqual(0, filter.SelectedValues.Count);
            }

            // 
            Assert.AreEqual(1, output.Applicablefilters[0].Values.Count);
            Assert.AreEqual("K", output.Applicablefilters[0].Values[0]);
        }
    }
}
