using FilesEditor.Tests.Constants;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;


namespace FilesEditor.Tests
{
    [TestClass]
    public class BaseTest
    {
        private const string TEST_FILES_FOLDER = "TestFiles";


        [TestInitialize]
        public void TestInitialize()
        {
            var fileDaCancellarePrimaDeiTest = new List<string>
                {
                TestPaths.OUTPUT_DEBUGFILE,
                // potenziali file di output generati dai test
                TestPaths.OUTPUT_FOLDER + "\\Engineering KPI_Direzione.pptx",
                TestPaths.OUTPUT_FOLDER + "\\Engineering KPI_Sales.pptx",
                TestPaths.OUTPUT_FOLDER + "\\Engineering KPI_Purchase.pptx",
                };
            foreach (var filename in fileDaCancellarePrimaDeiTest)
            {
                DeleleFileInTheTestFilesFolder(filename);
            }
        }

        private void DeleleFileInTheTestFilesFolder(string fileName)
        {
            var filePath = Path.Combine(TestFileFolderPath, fileName);
            if (File.Exists(filePath))
            {
                try
                {
                    File.Delete(filePath);
                }
                catch
                {
                    throw new Exception($"Impossibile cancellare il file {fileName}");
                }
            }
        }


        [TestCleanup]
        public void TestCleanup()
        { }

        public static string BinFolderPath
        {
            get
            {
                return  Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";
            }
        }

        public static string TestFileFolderPath
        {
            get
            {
                return Path.Combine(BinFolderPath, TEST_FILES_FOLDER);
            }
        }
    }
}