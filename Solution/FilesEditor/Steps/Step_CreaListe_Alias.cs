using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
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
            var ePPlusHelper = GetHelperForExistingFile(Context.DataSourceFilePath, FileTypes.DataSource);

            Context.AliasDefinitions_BusinessTMP = readAliasFromWorksheet(ePPlusHelper, WorksheetNames.DATA_SOURCE_ALIAS_BUSINESS_TMP);
            Context.AliasDefinitions_Categoria = readAliasFromWorksheet(ePPlusHelper, WorksheetNames.DATA_SOURCE_ALIAS_BUSINESS_CATEGORIA);
        }

        private List<AliasDefinition> readAliasFromWorksheet(EPPlusHelper ePPlusHelper, string worksheetName)
        {
            //  var worksheetName = WorksheetNames.DATA_SOURCE_ALIAS_BUSINESS_TMP;
            ThrowExpetionsForMissingWorksheet(ePPlusHelper, worksheetName, FileTypes.DataSource);

            var firstRow = Context.Configurazione.DATASOURCE_ALIAS_WORKSHEETS_FIRST_DATA_ROW;
            var lastRow = ePPlusHelper.GetLastUsedRowForColumn(worksheetName, firstRow, Context.Configurazione.DATASOURCE_ALIAS_WORKSHEETS_NEW_VALUES_COL);

            // mi posizione sulla riga precedente a quella da cui partire
            var currentRowNumber = firstRow - 1;

            List<AliasDefinition> aliasDefinitions = new List<AliasDefinition>();

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



                //// Verifico che il nome aliasMachineNameFrom (da usarsi come from nell'alias) non corrisponda al nome di una macchina
                //if (Context.AnagraficaMacchine.ContainsKey(aliasMachineNameFrom.ToLower()))
                //{
                //    throw new ManagedException(
                //        tipologiaErrore: TipologiaErrori.InvalidData,
                //        tipologiaPath: TipologiaPaths.AliasesFile,
                //        worksheetName: worksheetName,
                //        path: filePath,
                //        rigaCella: currentRowNumber,
                //        colonnaCella: Context.Configurazione.Aliases_MachineAliases_NomeFrom_Column,
                //        dato: aliasMachineNameFrom,
                //        nomeDatoConErrore: NomiDatoConErrore.AliasMachineNameFrom,
                //        messaggioPerUtente: $"The name '{aliasMachineNameFrom}' cannot be used as an alias because it already corresponds to the name of a machine."
                //        );
                //}


                //// cerco un'eventuale macchina che abbia già questo alias
                //var macchinaConStessoAlias = Context.AnagraficaMacchine.Values.FirstOrDefault(_ => _.HasThisAsNameOrAlias(aliasMachineNameFrom));
                //if (macchinaConStessoAlias != null && !macchinaConStessoAlias.Name.Equals(aliasMachineNameTo))
                //{
                //    // una macchina con lo stesso alias ha un "To" diverso, questo significa che si sta tendando di usare il From su due macchine diverse (un alias ambiguo)
                //    throw new ManagedException(
                //        tipologiaErrore: TipologiaErrori.InvalidData,
                //        tipologiaPath: TipologiaPaths.AliasesFile,
                //        worksheetName: worksheetName,
                //        path: filePath,
                //        rigaCella: currentRowNumber,
                //        colonnaCella: Context.Configurazione.Aliases_MachineAliases_NomeFrom_Column,
                //        dato: aliasMachineNameFrom,
                //        nomeDatoConErrore: NomiDatoConErrore.AliasMachineNameFrom,
                //        messaggioPerUtente: $"The alias '{aliasMachineNameFrom}' cannot be associated with the machine '{aliasMachineNameTo}' because it is already associated with the machine '{macchinaConStessoAlias.Name}'."
                //        );
                //}

                //// Cerco casi di alias già dichiarati
                //var aliasGiaEsistente = Context.AliasMacchine.FirstOrDefault(_ =>
                //                            _.MachineNameFrom.Equals(aliasMachineNameFrom, StringComparison.InvariantCultureIgnoreCase) &&
                //                            _.MachineNameTo.Equals(aliasMachineNameTo, StringComparison.InvariantCultureIgnoreCase));
                //if (aliasGiaEsistente != null)
                //{
                //    AddWarning($"The alias '{aliasMachineNameFrom}'-->'{aliasMachineNameTo}' is declared twice. Alias 1 file: '{fileName}' row number: {currentRowNumber} - Alias 2) file: '{aliasGiaEsistente.FileName}' row number: {aliasGiaEsistente.RowNumber}");
                //    continue;
                //}

                //// cerco la macchina destinazione (To) dell'associazione dell'alias (from --> to)
                //if (!Context.AnagraficaMacchine.ContainsKey(aliasMachineNameTo.ToLower()))
                //{
                //    AddWarning($"The alias '{aliasMachineNameFrom}'-->'{aliasMachineNameTo}' found into the file '{fileName}' at row number: {currentRowNumber} is invalid as there is no machine with the name '{aliasMachineNameTo}'.");
                //    continue;
                //}

                //// Aggiungo l'alias alla macchina corrispondente
                //Context.AnagraficaMacchine[aliasMachineNameTo.ToLower()].AggiungiAlias(aliasMachineNameFrom);
                //// Aggiungo l'alias alla lsita generale delli alases
                //Context.AliasMacchine.Add(new AliasMacchina(aliasMachineNameFrom, aliasMachineNameTo, fileName, currentRowNumber));


            }

            Context.DebugInfoLogger.LogAlias(aliasDefinitions, worksheetName);

            return aliasDefinitions;
        }

        private void creaLista_Alias_Categoria()
        {

        }
    }
}