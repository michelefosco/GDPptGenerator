using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Reflection;

namespace EPPlusExtensions.Tests
{
    [TestClass]
    public class BaseTest
    {
        private EPPlusHelper _excelHelper;
        private const string TEST_FILES_FOLDER = "TestFiles";

        [TestInitialize]
        public void TestInitialize()
        {
            _excelHelper = new EPPlusHelper();
        }

        [TestCleanup]
        public void TestCleanup()
        {
            _excelHelper = null;
        }


        public EPPlusHelper ExcelHelper
        {
            get
            {
                //if (_excelHelper == null)
                //    _excelHelper = new EPPlusHelper();
                return _excelHelper;
            }
        }


        public string TestFileFolderPath
        {
            get
            {
                var binFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";
                return Path.Combine(binFolder, TEST_FILES_FOLDER);
            }
        }
    }
}