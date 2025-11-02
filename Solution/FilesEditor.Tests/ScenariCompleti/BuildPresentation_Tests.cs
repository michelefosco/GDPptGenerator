using FilesEditor.Entities;
using FilesEditor.Enums;
using FilesEditor.Tests.Constants;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FilesEditor.Tests
{
    [TestClass]
    public class BuildPresentation_Tests : BaseTest
    {
        [TestMethod]
        public void Scenario_OK_01()
        {
            // properties base class
            string powerPointTemplateFilePath = Path.Combine(BinFolderPath, FilesEditor.Constants.FileNames.POWERPOINT_TEMPLATE_FILENAME);

            string dataSourceFolder = Path.Combine(TestFileFolderPath, TestPaths.DATASOURCE_FOLDER);
            string dataSourceFilePath = Path.Combine(dataSourceFolder, FilesEditor.Constants.FileNames.DATASOURCE_FILENAME);

            string destinationFolder = Path.Combine(TestFileFolderPath, TestPaths.OUTPUT_FOLDER);
            string tmpFolder = Path.Combine(destinationFolder, TestPaths.TMP_FOLDER);
            string debugFilePath = Path.Combine(destinationFolder, TestPaths.OUTPUT_DEBUGFILE);

            // properties specifiche di questo oggetto di input
            string fileBudgetPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_BUDGET_FILE);
            string fileForecastPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_FORECAST_FILE);
            string fileSuperDettagliPath = Path.Combine(TestFileFolderPath, TestPaths.INPUT_SUPERDETTAGLI_FILE);
            string fileRunRatePat = Path.Combine(TestFileFolderPath, TestPaths.INPUT_RUNRATE_FILE);
            bool appendCurrentYear_FileSuperDettagli = false;
            DateTime periodDate = DateTime.Today;

            //todo: settare filtri utili e valutarte gli effetti
            //var applicablefilters = new List<InputDataFilters_Item>
            //{
            //    new InputDataFilters_Item {
            //        Table = InputDataFilters_Tables.FORECAST,
            //        FieldName = "Field 1",
            //        SelectedValues = new List<string>{ "Valore 1", "Valore 2" }
            //    },
            //    new InputDataFilters_Item {
            //        Table = InputDataFilters_Tables.BUDGET,
            //        FieldName = "Field 2",
            //        SelectedValues = new List<string>{ "Valore 1", "Valore 2" }
            //    }
            //};
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
                    powerPointTemplateFilePath: powerPointTemplateFilePath,
                    appendCurrentYear_FileSuperDettagli: appendCurrentYear_FileSuperDettagli,
                    periodDate: periodDate,
                    applicablefilters: applicablefilters
                    );
            var output = Editor.BuildPresentation(input);



            int numeroRigheBudget = 34;
            int numeroRigheForecast = 34;
            int numeroRigheSuperdettagli = 1000;
            int numeroRigheRunRate = 1;
            //
            int numeroFilesFotoInTmpFolder = 14;
            int numeroWarnings = 0;
            int numeroPresentazioniGenerate = 3;

            CheckResults(dataSourceFilePath, tmpFolder, output, numeroRigheBudget, numeroRigheForecast, numeroRigheSuperdettagli, numeroRigheRunRate, numeroFilesFotoInTmpFolder, numeroWarnings, numeroPresentazioniGenerate);
        }

        private static void CheckResults(string dataSourceFilePath, string tmpFolder, Entities.MethodsArgs.BuildPresentationOutput output, int numeroRigheBudget, int numeroRigheForecast, int numeroRigheSuperdettagli, int numeroRigheRunRate, int numeroFilesFotoInTmpFolder, int numeroWarnings, int numeroPresentazioniGenerate)
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

            // numero di immagini generate su file system
            var filesInTmpFolder = Directory.GetFiles(tmpFolder);
            Assert.AreEqual(numeroFilesFotoInTmpFolder, filesInTmpFolder.Length);

            var ePPlusHelper = new EPPlusExtensions.EPPlusHelper();
            ePPlusHelper.Open(dataSourceFilePath);



            const int offSetIntestazioneBudget = 2;
            Assert.AreEqual(offSetIntestazioneBudget + numeroRigheBudget, ePPlusHelper.GetFirstEmptyRow(FilesEditor.Constants.WorksheetNames.DATASOURCE_BUDGET_DATA, 1, 1) - 1);
            //
            const int offSetIntestazioneForecast = 2;
            Assert.AreEqual(offSetIntestazioneForecast + numeroRigheForecast, ePPlusHelper.GetFirstEmptyRow(FilesEditor.Constants.WorksheetNames.DATASOURCE_FORECAST_DATA, 1, 1) - 1);
            //
            const int offSetIntestazioneRunRate = 1;
            Assert.AreEqual(offSetIntestazioneRunRate + numeroRigheRunRate, ePPlusHelper.GetFirstEmptyRow(FilesEditor.Constants.WorksheetNames.DATASOURCE_RUN_RATE_DATA, 1, 1) - 1);
            //
            const int offSetIntestazioneSuperdettagli = 2;
            Assert.AreEqual(offSetIntestazioneSuperdettagli + numeroRigheSuperdettagli, ePPlusHelper.GetFirstEmptyRow(FilesEditor.Constants.WorksheetNames.DATASOURCE_SUPERDETTAGLI_DATA, 2, 1) - 1);
        }
    }
}
