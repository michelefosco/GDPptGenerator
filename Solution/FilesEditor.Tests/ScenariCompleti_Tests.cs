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
            string fileRanRatePat = Path.Combine(TestFileFolderPath, TestPaths.INPUT_RANRATE_FILE);


            var validaSourceFilesInput = new ValidaSourceFilesInput(
                    destinationFolder: destinationFolder,
                    templatesFolder: templatesFolder,
                    fileBudgetPath: fileBudgetPath,
                    fileForecastPath: fileForecastPath,
                    fileSuperDettagliPath: fileSuperDettagliPath,
                    fileRanRatePath: fileRanRatePat
                    );

            var validaSourceFilesOutput = Editor.ValidaSourceFiles(validaSourceFilesInput);

            Assert.IsNotNull(validaSourceFilesOutput);
            Assert.AreEqual(EsitiFinali.Success, validaSourceFilesOutput.Esito);
            Assert.IsNull(validaSourceFilesOutput.ManagedException);
            Assert.IsNotNull(validaSourceFilesOutput.UserOptions);
            //
            Assert.IsNotNull(validaSourceFilesOutput.UserOptions.Applicablefilters);
            Assert.AreEqual(8, validaSourceFilesOutput.UserOptions.Applicablefilters.Count);
            //
            Assert.IsNotNull(validaSourceFilesOutput.UserOptions.SildeToGenerate);
            Assert.AreEqual(18, validaSourceFilesOutput.UserOptions.SildeToGenerate.Count);
        }
    }
}
