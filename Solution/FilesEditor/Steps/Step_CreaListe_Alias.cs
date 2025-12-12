using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FilesEditor.Steps
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_CreaListe_Alias : StepBase
    {
        public override string StepName => "Step_CreaListe_Alias";

        internal override void BeforeTask()
        {
            Context.DebugInfoLogger.LogStepContext(StepName, Context);
        }

        internal override void ManageInfoAboutPerformedStepTask(TimeSpan timeSpent)
        {
            Context.DebugInfoLogger.LogPerformance(StepName, timeSpent);
        }

        internal override void AfterTask()
        {
            Context.DebugInfoLogger.LogStepContext(StepName, Context);
        }
        public Step_CreaListe_Alias(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            CreaListe_Alias();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void CreaListe_Alias()
        {
            FillAliasesFromWorksheet(Context.DataSourceEPPlusHelper, WorksheetNames.DATASOURCE_ALIAS_BUSINESS, Context.AliasDefinitions_Business);
            FillAliasesFromWorksheet(Context.DataSourceEPPlusHelper, WorksheetNames.DATASOURCE_ALIAS_BUSINESS_CATEGORIA, Context.AliasDefinitions_Categoria);
        }

        private void FillAliasesFromWorksheet(EPPlusHelper ePPlusHelper, string worksheetName, List<AliasDefinition> aliases)
        {
            EPPlusHelperUtilities.ThrowExpetionsForMissingWorksheet(ePPlusHelper, worksheetName, FileTypes.DataSource);

            var firstRow = Context.Configurazione.DATASOURCE_ALIAS_WORKSHEETS_FIRST_DATA_ROW;
            var lastRow = ePPlusHelper.GetLastUsedRowForColumn(worksheetName, firstRow, Context.Configurazione.DATASOURCE_ALIAS_WORKSHEETS_NEW_VALUES_COL);

            // mi posiziono sulla riga precedente a quella da cui partire
            var currentRowNumber = firstRow - 1;

            // se riga corrente diventa = a lastRow devo fermarmi, in quanto verrebbe incrementata di 1 e si andrebbe oltre l'ultima riga
            while (currentRowNumber < lastRow)
            {
                // mi posiziono sulla prossima riga da leggere
                currentRowNumber++;

                // leggo i valori da cercare e quelli da usarsi come sostituti
                var rawValues = ePPlusHelper.GetString(worksheetName, currentRowNumber, Context.Configurazione.DATASOURCE_ALIAS_WORKSHEETS_RAW_VALUES_COL);
                var newValue = ePPlusHelper.GetString(worksheetName, currentRowNumber, Context.Configurazione.DATASOURCE_ALIAS_WORKSHEETS_NEW_VALUES_COL);

                // ignoro le righe vuote (entrambi i campi a null)
                if (string.IsNullOrEmpty(rawValues) && string.IsNullOrEmpty(newValue))
                { continue; }

                // Loggo un warning per ogni alias scritto male
                if (string.IsNullOrEmpty(rawValues) || string.IsNullOrEmpty(newValue))
                {
                    Context.AddWarning($"The alias declared in the worksheet '{worksheetName}' at line: {currentRowNumber} is incomplete.");
                    continue;
                }

                // splitto i raw values (separati da ,)
                var rawValuesSplittati = rawValues.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(_ => _.Trim()).Where(_ => !string.IsNullOrWhiteSpace(_)).ToList();
                foreach (var rawValueSplittato in rawValuesSplittati)
                {
                    if (aliases.Any(_ => _.RawValue.Equals(rawValueSplittato, StringComparison.InvariantCultureIgnoreCase)))
                    {
                        throw new ManagedException(
                                    filePath: ePPlusHelper.FilePathInUse,
                                    fileType: FileTypes.DataSource,
                                    //
                                    worksheetName: worksheetName,
                                    cellRow: currentRowNumber,
                                    cellColumn: null,
                                    valueHeader: ValueHeaders.None,
                                    value: rawValueSplittato,
                                    //
                                    errorType: ErrorTypes.DuplicateValue,
                                    userMessage: $"The alias '{rawValueSplittato}' is declared more than once in the worksheet '{worksheetName}'.");
                    }
                    aliases.Add(new AliasDefinition(rawValue: rawValueSplittato, newValue: newValue));
                }
            }

            Context.DebugInfoLogger.LogAlias(aliases, worksheetName);
        }
    }
}