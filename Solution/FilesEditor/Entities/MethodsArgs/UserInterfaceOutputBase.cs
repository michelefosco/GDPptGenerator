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

        internal UserInterfaceOutputBase(StepContext context)
        {
            Esito = context.Esito;
            DebugFilePath = context.DebugFilePath;
            Warnings = context.Warnings;
            ElapsedTime = context.ElapsedTime;
        }
        internal UserInterfaceOutputBase(ManagedException managedException)
        {
            Esito = EsitiFinali.Failure;
            ManagedException = managedException;
        }
    }
}