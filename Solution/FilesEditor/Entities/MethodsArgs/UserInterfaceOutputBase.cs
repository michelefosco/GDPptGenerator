using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using System;
using System.Collections.Generic;

namespace FilesEditor.Entities.MethodsArgs
{
    public class UserInterfaceOutputBase
    {
        public ManagedException ManagedException { get; private set; }
        public EsitiFinali Esito { get; private set; }
        public string DebugFilePath { get; set; }
        public List<string> Warnings { get; set; }
        public TimeSpan ElapsedTime { get; set; }

        internal UserInterfaceOutputBase(StepContext context, ManagedException managedException = null)
        {
            if (managedException != null)
            {
                Esito = EsitiFinali.Failure;
                ManagedException = managedException;
            }
            else
            {
                Esito = context.Esito;
            }

            DebugFilePath = context.DebugFilePath;
            Warnings = context.Warnings;
            ElapsedTime = context.ElapsedTime;
        }
    }
}