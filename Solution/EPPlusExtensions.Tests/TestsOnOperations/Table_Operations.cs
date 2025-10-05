using EPPlusExtensions.Tests.Constants;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace EPPlusExtensions.Tests.TestsOnOperations
{
    [TestClass]
    public class Table_Operations : BaseTest
    {
        [DataTestMethod]
        [DataRow(1, 0, true)]
        [DataRow(2, 1, true)]
        [DataRow(3, 5, true)]
        [DataRow(4, 9, true)]
        [DataRow(1, 13, false)]
        [DataRow(2, 13, false)]
        [DataRow(3, 13, false)]
        [DataRow(4, 13, false)]
        public void GetLastUsedRowForColumn(int colonna, int expectedLastUsedRowForColumn, bool ignoreBlanks)
        {
            //GIVEN
            var filePath = Path.Combine(TestFileFolderPath, InputfileNames.INPUT_FILE_001);
            ExcelHelper.Open(filePath);

            // WHEN

            var actualLastUsedRowForColumn = ExcelHelper.GetLastUsedRowForColumn("FunzioneGetLastUsedRowForColumn",1 , colonna, ignoreBlanks);

            // THEN            
            Assert.AreEqual(expectedLastUsedRowForColumn, actualLastUsedRowForColumn);
        }

        [DataTestMethod]
        [DataRow(1, -1, true)]
        [DataRow(2, 7, false)]
        [DataRow(3, 3, true)]
        [DataRow(3, 3, false)]
        [DataRow(4, 11, true)]
        [DataRow(5, 7, false)]
        public void GetFirstRowWithAnyValue(int colonna, int expectedLastUsedRowForColumn, bool ignoreBlanks)
        {
            //GIVEN
            var filePath = Path.Combine(TestFileFolderPath, InputfileNames.INPUT_FILE_001);
            ExcelHelper.Open(filePath);

            // WHEN

            var actualLastUsedRowForColumn = ExcelHelper.GetFirstRowWithAnyValue("FunzioneGetFirstRowWithAnyValue", 1, colonna, ignoreBlanks);

            // THEN            
            Assert.AreEqual(expectedLastUsedRowForColumn, actualLastUsedRowForColumn);
        }
        
    }
}