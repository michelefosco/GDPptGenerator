using FilesEditor.Entities;
using FilesEditor.Tests.Constants;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace FilesEditor.Tests
{
    [TestClass]
    public class ScenariCompleti_Tests: BaseTest
    {
        [TestMethod]
        public void Interazione_OK_01()
        {

            var templatesFolder_Path = Path.Combine(TestFileFolderPath, TestPaths.TEMPLATES_FOLDER);

            //var folderOutput_Path = Path.Combine(TestFileFolderPath, TestPaths.OUTPUT_FOLDER);


            //var fileOutput_Path = Path.Combine(TestFileFolderPath, TestPaths.OUTPUT_FILE);
            //var fileDebug_Path = Path.Combine(TestFileFolderPath, TestPaths.OUTPUT_DEBUGFILE);

            //var folderDatabase_Path = Path.Combine(TestFileFolderPath, TestPaths.KO_MacchineNomiDuplicatiStessoFile);
            //var folderVenduto_Path = Path.Combine(TestFileFolderPath, TestPaths.Venduto_DatiAdHoc);


            var validaSourceFilesInput = new ValidaSourceFilesInput( 
                templatesFolder: templatesFolder_Path
                );

            var validaSourceFilesOutput = Editor.ValidaSourceFiles(validaSourceFilesInput);

            //todo: inserire validazioni qui
            Assert.IsNotNull(validaSourceFilesOutput);
            Assert.IsNotNull(validaSourceFilesOutput.OpzioniUtente);


            //var createPresentationsInput = new CreatePresentationsInput(
            //    outputFolder: "Output",
            //    tmpFolder: "Tmp",
            //    templatesFolder: templatesFolder,
            //    fileDebug_FilePath: "fileDebug_FilePath"
            //    );
            //var createPresentationsOutput = Editor.CreatePresentations(createPresentationsInput);
        }
    }
}
