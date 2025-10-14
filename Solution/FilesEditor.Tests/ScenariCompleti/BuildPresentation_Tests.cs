using FilesEditor.Entities;
using FilesEditor.Enums;
using FilesEditor.Tests.Constants;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;

namespace FilesEditor.Tests
{
    [TestClass]
    public class BuildPresentation_Tests : BaseTest
    {
        [TestMethod]
        public void Scenario_OK_01()
        {
            // properties base class
            string sourceFilesFolderPath = Path.Combine(TestFileFolderPath, TestPaths.SOURCEFILES_FOLDER);
            string destinationFolder = Path.Combine(TestFileFolderPath, TestPaths.OUTPUT_FOLDER);
            string tmpFolder = Path.Combine(destinationFolder, TestPaths.TMP_FOLDER);
            string fileDebugPath = Path.Combine(destinationFolder, TestPaths.OUTPUT_DEBUGFILE);

            // properties specifiche di questo oggetto di input
            string fileBudgetPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_BUDGET_FILE);
            string fileForecastPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_FORECAST_FILE);
            string fileSuperDettagliPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_SUPERDETTAGLI_FILE);
            string fileRunRatePat = Path.Combine(TestFileFolderPath, TestPaths.INPUT_RUNRATE_FILE);
            bool replaceAllData_FileSuperDettagli = true;
            DateTime periodDate = DateTime.Today;
            //todo utilizzare
            var applicablefilters = new List<InputDataFilters_Item>();

            var input = new Entities.MethodsArgs.BuildPresentationInput(
                    sourceFilesFolderPath: sourceFilesFolderPath,
                    destinationFolder: destinationFolder,
                    tmpFolder: tmpFolder,
                    fileDebugPath: fileDebugPath,
                    //
                    replaceAllData_FileSuperDettagli: replaceAllData_FileSuperDettagli,
                    periodDate: periodDate,
                    applicablefilters: applicablefilters
                    );
            var output = Editor.BuildPresentation(input);

            // test base
            Assert.IsNotNull(output);
            Assert.AreEqual(EsitiFinali.Success, output.Esito);
            Assert.IsNull(output.ManagedException);

            // test specifici dell'oggetto di output
            //
            //Assert.IsNotNull(buildPresentationOutput.UserOptions.Applicablefilters);
            //Assert.AreEqual(8, buildPresentationOutput.UserOptions.Applicablefilters.Count);

            //foreach (var filter in buildPresentationOutput.UserOptions.Applicablefilters)
            //{
            //    Assert.IsFalse(string.IsNullOrWhiteSpace(filter.FieldName));
            //    Assert.IsNotNull(filter.Values);
            //    Assert.IsTrue(filter.Values.Count > 0);
            //    Assert.AreEqual(0, filter.SelectedValues.Count);
            //}

            //// 
            //Assert.AreEqual(1, buildPresentationOutput.UserOptions.Applicablefilters[0].Values.Count);
            //Assert.AreEqual("K", buildPresentationOutput.UserOptions.Applicablefilters[0].Values[0]);           
            ////
            //Assert.IsNotNull(buildPresentationOutput.UserOptions.SildeToGenerate);
            //Assert.AreEqual(18, buildPresentationOutput.UserOptions.SildeToGenerate.Count);
        }
    }
}
