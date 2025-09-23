using EPPlusExtensions.Tests.Constants;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace EPPlusExtensions.Tests.TestsOnOperations
{
    [TestClass]
    public class Open_Operation : BaseTest
    {
        [TestMethod]
        public void Open_WithRightFilePath_ReturnsTrue()
        {
            // GIVEN
            var filePath = Path.Combine(TestFileFolderPath, InputfileNames.INPUT_FILE_001);

            // WHEN
            var openSuccess = ExcelHelper.Open(filePath);

            // THEN
            Assert.IsTrue(openSuccess);
        }


        [TestMethod]
        public void Create_WorksOK()
        {
            // GIVEN
            var filePath = Path.Combine(TestFileFolderPath, InputfileNames.OUTPUT_FILE_001);
            if (File.Exists(filePath))
                File.Delete(filePath);

            // WHEN
            var openSuccess = ExcelHelper.Create(filePath, "log");

            // THEN
            Assert.IsTrue(openSuccess);
        }


        [TestMethod]
        public void Open_WithWrongFilePath_ReturnsTrue()
        {
            // GIVEN
            var filePath = Path.Combine(TestFileFolderPath, "WrongName.xlsx");

            // WHEN
            var openSuccess = ExcelHelper.Open(filePath);

            // THEN
            Assert.IsFalse(openSuccess);
        }
    }
}