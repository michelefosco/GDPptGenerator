using FilesEditor.Entities;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using System;
using System.IO;

namespace FilesEditor.Steps
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_TmpFolder_Predisposizione : StepBase
    {
        internal override string StepName => "Step_TmpFolder_Predisposizione";

        public Step_TmpFolder_Predisposizione(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            PredisponiTmpFolder();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void PredisponiTmpFolder()
        {
            // Rimuovo la cartella se già esistente
            FilesAndDirectoriesUtilities.CancellaDirectorySeEsiste(Context.TmpFolder);

            // Creo la cartella
            Directory.CreateDirectory(Context.TmpFolder);
        }
    }
}