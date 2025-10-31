using FilesEditor.Entities;
using FilesEditor.Enums;

namespace FilesEditor.Steps.BuildPresentation
{
    internal class Step_SalvaFile_DataSource: StepBase
    {
        public Step_SalvaFile_DataSource(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            Context.DebugInfoLogger.LogStepContext("Step_SalvaDataSource", Context);

            Context.EpplusHelperDataSource.Save();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }
    }
}