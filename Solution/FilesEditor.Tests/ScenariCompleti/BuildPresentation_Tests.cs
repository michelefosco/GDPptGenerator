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
            string dataSourceFolder = Path.Combine(TestFileFolderPath, TestPaths.SOURCEFILES_FOLDER);
            string dataSourceFilePath = Path.Combine(dataSourceFolder, FilesEditor.Constants.FileNames.DATASOURCE_FILENAME);
            string destinationFolder = Path.Combine(TestFileFolderPath, TestPaths.OUTPUT_FOLDER);
            string tmpFolder = Path.Combine(destinationFolder, TestPaths.TMP_FOLDER);
            string debugFilePath = Path.Combine(destinationFolder, TestPaths.OUTPUT_DEBUGFILE);

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
                    dataSourceFilePath: dataSourceFilePath,
                    destinationFolder: destinationFolder,
                    tmpFolder: tmpFolder,
                    debugFilePath: debugFilePath,
                    //
                    fileBudgetPath: fileBudgetPath,
                    fileForecastPath: fileForecastPath,
                    fileSuperDettagliPath: fileSuperDettagliPath,
                    fileRunRatePath: fileRunRatePat,
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

            // numero file di output generati
            Assert.AreEqual(3, output.OutputFilePathLists.Count);

            // numero di warnings sollevati
            Assert.AreEqual(2, output.Warnings.Count);

            // numero di immagini generate su file system
            var filesInTmpFolder = Directory.GetFiles(tmpFolder);
            Assert.AreEqual (14, filesInTmpFolder.Length);
        }
    }
}
