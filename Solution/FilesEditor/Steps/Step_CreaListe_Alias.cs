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
    internal class Step_CreaListe_Alias : StepBase
    {
        public Step_CreaListe_Alias(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            Context.DebugInfoLogger.LogStepContext("Step_CreaListe_Alias", Context);
            creaListe_Alias();
            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void creaListe_Alias()
        {
           // var ePPlusHelper = EPPlusHelperUtilities.GetEPPlusHelperForExistingFile(Context.DataSourceFilePath, FileTypes.DataSource);

            Context.AliasDefinitions_BusinessTMP = readAliasFromWorksheet(Context.ePPlusHelperDataSource, WorksheetNames.DATASOURCE_ALIAS_BUSINESS_TMP);
            Context.AliasDefinitions_Categoria = readAliasFromWorksheet(Context.ePPlusHelperDataSource, WorksheetNames.DATASOURCE_ALIAS_BUSINESS_CATEGORIA);
        }

        private List<AliasDefinition> readAliasFromWorksheet(EPPlusHelper ePPlusHelper, string worksheetName)
        {
            //  var worksheetName = WorksheetNames.DATASOURCE_ALIAS_BUSINESS_TMP;
            EPPlusHelperUtilities.ThrowExpetionsForMissingWorksheet(ePPlusHelper, worksheetName, FileTypes.DataSource);

            var firstRow = Context.Configurazione.DATASOURCE_ALIAS_WORKSHEETS_FIRST_DATA_ROW;
            var lastRow = ePPlusHelper.GetLastUsedRowForColumn(worksheetName, firstRow, Context.Configurazione.DATASOURCE_ALIAS_WORKSHEETS_NEW_VALUES_COL);

            // mi posizione sulla riga precedente a quella da cui partire
            var currentRowNumber = firstRow - 1;

            var aliasDefinitions = new List<AliasDefinition>();

            // se riga corrente diventa = a lastRow devo fermarmi, in quanto verrebbe incrementata di 1 e si andrebbe oltre l'ultima riga
            while (currentRowNumber < lastRow)
            {
                // mi posiziono sulla prossima riga da leggere
                currentRowNumber++;

                var rawValues = ePPlusHelper.GetString(worksheetName, currentRowNumber, Context.Configurazione.DATASOURCE_ALIAS_WORKSHEETS_RAW_VALUES_COL);
                var newValue = ePPlusHelper.GetString(worksheetName, currentRowNumber, Context.Configurazione.DATASOURCE_ALIAS_WORKSHEETS_NEW_VALUES_COL);

                // ignoro le righe vuote (entrambi i campi a null)
                if (string.IsNullOrEmpty(rawValues) && string.IsNullOrEmpty(newValue))
                { continue; }


                if (string.IsNullOrEmpty(rawValues) || string.IsNullOrEmpty(newValue))
                {
                    AddWarning($"The alias declared in the worksheet '{worksheetName}' at line: {currentRowNumber} is incomplete.");
                    continue;
                }

                // splitto i raw values (separati da ,)
                var rawValuesSplittati = rawValues.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(_ => _.Trim()).Where(_ => !string.IsNullOrWhiteSpace(_)).ToList();
                foreach (var rawValueSplittato in rawValuesSplittati)
                {
                    if (aliasDefinitions.Any(_ => _.RawValue.Equals(rawValueSplittato, StringComparison.InvariantCultureIgnoreCase)))
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

                    aliasDefinitions.Add(new AliasDefinition(
                        rawValue: rawValueSplittato,
                        newValue: newValue));
                }
            }

            Context.DebugInfoLogger.LogAlias(aliasDefinitions, worksheetName);

            return aliasDefinitions;
        }
    }
}