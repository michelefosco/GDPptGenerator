using EPPlusExtensions.Tests.Constants;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace EPPlusExtensions.Tests.TestsOnOperations
{
    [TestClass]
    public class Read_Operations : BaseTest
    {
        [TestMethod]
        public void GetWorksheetNames_WorksOK()
        {
            // GIVEN
            var filePath = Path.Combine(TestFileFolderPath, SourceFileNames.INPUT_FILE_001);
            ExcelHelper.Open(filePath);

            // WHEN
            var listName = ExcelHelper.GetWorksheetNames();

            // THEN
            Assert.AreEqual(7, listName.Count);
            Assert.IsTrue(listName.Contains("FOGLIO PRIMO"));
            Assert.IsTrue(listName.Contains("FORMULE"));
            Assert.IsTrue(listName.Contains("DATIPERPIVOT"));
            Assert.IsTrue(listName.Contains("SHEET3"));
            Assert.IsTrue(listName.Contains("FunzioneGetLastUsedRowForColumn".ToUpper()));
            Assert.IsTrue(listName.Contains("FunzioneGetFirstRowWithAnyValue".ToUpper()));
        }

        [TestMethod]
        public void GetRowsLimit_WorksOK()
        {
            // GIVEN
            var filePath = Path.Combine(TestFileFolderPath, SourceFileNames.INPUT_FILE_002);
            ExcelHelper.Open(filePath);

            // WHEN
            var actualcount = ExcelHelper.GetRowsLimit("FBL3N_act S510760056");

            // THEN
            Assert.AreEqual(29, actualcount);
        }


        [DataTestMethod]
        [DataRow(1, 2, false)]
        [DataRow(2, 2, false)]
        [DataRow(3, 2, false)]
        [DataRow(4, 2, false)]
        [DataRow(5, 2, false)]
        [DataRow(6, 2, false)]
        [DataRow(7, 2, true)]
        public void IsFormula(int row, int column, bool expectedIsFormula)
        {
            // GIVEN
            var filePath = Path.Combine(TestFileFolderPath, SourceFileNames.INPUT_FILE_001);
            ExcelHelper.Open(filePath);

            // WHEN
            var actualIsFormula = ExcelHelper.IsFormula("FOGLIO PRIMO", row, column);

            // THEN
            Assert.AreEqual(expectedIsFormula, actualIsFormula);
        }

        [DataTestMethod]
        [DataRow("FoglioSenzaCelle", false)]
        [DataRow("DatiPerPivot", true)]
        [DataRow("Foglio primo", true)]
        public void WorksheetExists(string worksheetName, bool expectedHasWorksheetCells)
        {
            // GIVEN
            var filePath = Path.Combine(TestFileFolderPath, SourceFileNames.INPUT_FILE_001);
            ExcelHelper.Open(filePath);

            // WHEN
            var actualHasWorksheetCells = ExcelHelper.WorksheetExists(worksheetName);

            // THEN
            Assert.AreEqual(expectedHasWorksheetCells, actualHasWorksheetCells);
        }

    }
}