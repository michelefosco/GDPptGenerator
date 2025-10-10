using FilesEditor.Entities;
using FilesEditor.Enums;
using FilesEditor.Tests.Constants;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace FilesEditor.Tests
{
    [TestClass]
    public class ScenariCompleti_Tests : BaseTest
    {
        [TestMethod]
        public void Interazione_OK_01()
        {
            string destinationFolder = Path.Combine(TestFileFolderPath, TestPaths.OUTPUT_FOLDER); ;
            string templatesFolder = Path.Combine(TestFileFolderPath, TestPaths.TEMPLATES_FOLDER);
            //
            string fileBudgetPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_BUDGET_FILE);
            string fileForecastPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_FORECAST_FILE);
            string fileSuperDettagliPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_SUPERDETTAGLI_FILE);
            string fileRunRatePat = Path.Combine(TestFileFolderPath, TestPaths.INPUT_RUNRATE_FILE);


            var validaSourceFilesInput = new ValidaSourceFilesInput(
                    destinationFolder: destinationFolder,
                    templatesFolder: templatesFolder,
                    fileBudgetPath: fileBudgetPath,
                    fileForecastPath: fileForecastPath,
                    fileSuperDettagliPath: fileSuperDettagliPath,
                    fileRunRatePath: fileRunRatePat
                    );

            var validaSourceFilesOutput = Editor.ValidaSourceFiles(validaSourceFilesInput);

            Assert.IsNotNull(validaSourceFilesOutput);
            Assert.AreEqual(EsitiFinali.Success, validaSourceFilesOutput.Esito);
            Assert.IsNull(validaSourceFilesOutput.ManagedException);
            Assert.IsNotNull(validaSourceFilesOutput.UserOptions);
            //
            Assert.IsNotNull(validaSourceFilesOutput.UserOptions.Applicablefilters);
            Assert.AreEqual(8, validaSourceFilesOutput.UserOptions.Applicablefilters.Count);

            foreach (var filter in validaSourceFilesOutput.UserOptions.Applicablefilters)
            {
                Assert.IsFalse(string.IsNullOrWhiteSpace(filter.FieldName));
                Assert.IsNotNull(filter.Values);
                Assert.IsTrue(filter.Values.Count > 0);
                Assert.AreEqual(0, filter.SelectedValues.Count);
            }

            // 
            Assert.AreEqual(1, validaSourceFilesOutput.UserOptions.Applicablefilters[0].Values.Count);
            Assert.AreEqual("K", validaSourceFilesOutput.UserOptions.Applicablefilters[0].Values[0]);           
            //
            Assert.IsNotNull(validaSourceFilesOutput.UserOptions.SildeToGenerate);
            Assert.AreEqual(18, validaSourceFilesOutput.UserOptions.SildeToGenerate.Count);
        }
    }
}
