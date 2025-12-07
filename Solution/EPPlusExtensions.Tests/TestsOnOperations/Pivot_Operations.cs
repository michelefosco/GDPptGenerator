using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EPPlusExtensions.Tests.TestsOnOperations
{
    [TestClass]
    public class Pivot_Operations : BaseTest
    {
        //[TestMethod]
        //public void CreatePivotFromTable_WorksCorrectly()
        //{
        //    // GIVEN
        //    var originFilePath = Path.Combine(TestFileFolderPath, InputfileNames.INPUT_FILE_001);
        //    var savedAsFilePath = Path.Combine(TestFileFolderPath, "OutputTestFile_WithPivotAdded.xlsx");

        //    if (File.Exists(savedAsFilePath))
        //        File.Delete(savedAsFilePath);

        //    Assert.IsFalse(File.Exists(savedAsFilePath));

        //    // WHEN
        //    ExcelHelper.CreatePivotFromTable(originFilePath, "DatiPerPivot", savedAsFilePath);

        //    // THEN
        //    var actualCount = ExcelHelper.GetNumberOfFormulasInFile(savedAsFilePath, 1);
        //    Assert.IsTrue(File.Exists(savedAsFilePath));
        //}

        //[TestMethod]
        //public void RefreshAllPivots()
        //{
        //    // GIVEN
        //    var originFilePath = Path.Combine(TestFileFolderPath, InputfileNames.INPUT_FILE_003);
        //    var savedAsFilePath = Path.Combine(TestFileFolderPath, "OutputTestFile_RefreshedPivots.xlsx");

        //    if (File.Exists(savedAsFilePath))
        //        File.Delete(savedAsFilePath);

        //    Assert.IsFalse(File.Exists(savedAsFilePath));

        //    // WHEN
        //    ExcelHelper.Open(originFilePath);

        //    ExcelHelper.SetRefreshOnLoadForAllPivotTables();

        //    ExcelHelper.SaveAs(savedAsFilePath);

        //    // THEN
        //    Assert.IsTrue(File.Exists(savedAsFilePath));
        //}
    }
}