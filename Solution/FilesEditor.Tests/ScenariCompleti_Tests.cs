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
            string destinationFolder = "?";
            string templatesFolder = Path.Combine(TestFileFolderPath, TestPaths.TEMPLATES_FOLDER);
            string fileBudgetPath = "?";
            string fileForecastPath = "?";
            string fileSuperDettagliPath = "?";
            string fileRanRatePat = "?";


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
