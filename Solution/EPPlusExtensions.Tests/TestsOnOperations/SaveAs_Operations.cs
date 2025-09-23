using EPPlusExtensions.Tests.Constants;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace EPPlusExtensions.Tests.TestsOnOperations
{
    [TestClass]
    public class SaveAs_Operations : BaseTest
    {

        [TestMethod]
        public void SaveAs_WorksOK()
        {
            // GIVEN
            var originFilePath = Path.Combine(TestFileFolderPath, InputfileNames.INPUT_FILE_001);
            var savedAsFilePath = Path.Combine(TestFileFolderPath, "NewNameFor_" + InputfileNames.INPUT_FILE_001);

            // This makes sure the saveAs file does not exist
            if (File.Exists(savedAsFilePath)) { File.Delete(savedAsFilePath); }
            Assert.IsFalse(File.Exists(savedAsFilePath));

            // WHEN
            ExcelHelper.Open(originFilePath);
            ExcelHelper.SaveAs(savedAsFilePath);

            // THEN
            Assert.IsTrue(File.Exists(savedAsFilePath));
        }
    }
}