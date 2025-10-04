using FilesEditor.Entities;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace FilesEditor.Tests
{
    [TestClass]
    public class ScenariCompleti
    {
        [TestMethod]
        public void Interazione_OK_01()
        {
            var templatesFolder = "";

            var validaSourceFilesInput = new ValidaSourceFilesInput( 
                templateFolder: templatesFolder
                );

            var validaSourceFilesOutput = Info.ValidaSourceFiles(validaSourceFilesInput);

            //todo: inserire validazioni qui
            Assert.IsNotNull(validaSourceFilesOutput);
            Assert.IsNotNull(validaSourceFilesOutput.OpzioniUtente);


            var createPresentationsInput = new CreatePresentationsInput(
                outputFolder: "Output",
                tmpFolder: "Tmp",
                templatesFolder: templatesFolder,
                fileDebug_FilePath: "fileDebug_FilePath"
                );
            var createPresentationsOutput = Editor.CreatePresentations(createPresentationsInput);
        }
    }
}
