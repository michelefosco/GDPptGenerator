using FilesEditor.Entities;
using FilesEditor.Helpers;
using FilesEditor.Entities.MethodsArgs;
using System.IO;


namespace FilesEditor.Steps.ValidateSourceFiles
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_Start_FileDebugHelper : ValidateSourceFiles_StepBase
    {

        public Step_Start_FileDebugHelper(StepContext context) : base(context)
        { }

        internal override ValidaSourceFilesOutput DoSpecificTask()
        {
            inizializzaDebugInfoLogger();
            return null; // Step intermedio, non ritorna alcun esito
        }

        private void inizializzaDebugInfoLogger()
        {
            if (File.Exists(Context.FileDebugPath))
            {
                File.Delete(Context.FileDebugPath);
            }
            Context.DebugInfoLogger = new FileDebugHelper(Context.FileDebugPath, Context.Configurazione.AutoSaveDebugFile);
            //todo: considere se rimettere in fuzione questo
            //   Context.DebugInfoLogger.LogBuildPresentationInput(Context.BuildPresentationInput);
        }
    }
}