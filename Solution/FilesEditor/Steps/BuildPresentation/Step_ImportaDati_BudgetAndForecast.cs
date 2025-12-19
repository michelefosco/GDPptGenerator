using EPPlusExtensions;
using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_ImportaDati_BudgetAndForecast : StepBase
    {
        internal override string StepName => "Step_ImportaDati_BudgetAndForecast";

        public Step_ImportaDati_BudgetAndForecast(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            ImportaSourceFile(
                    sourceFileType: FileTypes.Budget,
                    sourceFileEPPlusHelper: Context.BudgetFileEPPlusHelper,
                    // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome
                    sourceWorksheetName: null, // WorksheetNames.SOURCEFILE_BUDGET_DATA,
                    souceHeadersRow: Context.Configurazione.SOURCE_FILES_BUDGET_HEADERS_ROW,
                    sourceHeadersFirstColumn: Context.Configurazione.SOURCE_FILES_BUDGET_HEADERS_FIRST_COL,
                    //
                    destWorksheetName: WorksheetNames.DATASOURCE_BUDGET_DATA,
                    destHeadersRow: Context.Configurazione.DATASOURCE_BUDGET_HEADERS_ROW,
                    destHeadersFirstColumn: Context.Configurazione.DATASOURCE_BUDGET_HEADERS_FIRST_COL
                );

            ImportaSourceFile(
                    sourceFileType: FileTypes.Forecast,
                    sourceFileEPPlusHelper: Context.ForecastFileEPPlusHelper,
                    // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome
                    sourceWorksheetName: null, // WorksheetNames.SOURCEFILE_FORECAST_DATA,
                    souceHeadersRow: Context.Configurazione.SOURCE_FILES_FORECAST_HEADERS_ROW,
                    sourceHeadersFirstColumn: Context.Configurazione.SOURCE_FILES_FORECAST_HEADERS_FIRST_COL,
                    //
                    destWorksheetName: WorksheetNames.DATASOURCE_FORECAST_DATA,
                    destHeadersRow: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_ROW,
                    destHeadersFirstColumn: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_FIRST_COL
                );

            //Context.BudgetFileEPPlusHelper.Close();
            //Context.ForecastFileEPPlusHelper.Close();

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void ImportaSourceFile(
                FileTypes sourceFileType,
                EPPlusHelper sourceFileEPPlusHelper,
                string sourceWorksheetName,
                int souceHeadersRow,
                int sourceHeadersFirstColumn,
                //
                string destWorksheetName,
                int destHeadersRow,
                int destHeadersFirstColumn
            )
        {
            if (sourceFileType != FileTypes.Budget && sourceFileType != FileTypes.Forecast)
            { throw new ArgumentOutOfRangeException(nameof(sourceFileType)); }


            #region WorkSheets sorgente e destinazione
            // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome
            var sourceWorksheet = sourceFileEPPlusHelper.ExcelPackage.Workbook.Worksheets[1];

            // Foglio destinazione
            var destWorksheet = Context.DataSourceEPPlusHelper.ExcelPackage.Workbook.Worksheets[destWorksheetName];
            #endregion

            // Variabili per il conteggio delle righe elaborate
            var infoRowsDestinazione = new InfoRows();
            infoRowsDestinazione.Iniziali = destWorksheet.Dimension.End.Row - destHeadersRow;

            #region Individuo gli eventuali filtri da applicare
            var inputDataFilters_Table = (sourceFileType == FileTypes.Budget)
                       ? InputDataFilters_Tables.BUDGET
                       : InputDataFilters_Tables.FORECAST;

            var filterBusiness = Context.ApplicableFilters.FirstOrDefault(_ => _.Table == inputDataFilters_Table
                                                                            && _.FieldName.Equals(Values.HEADER_BUSINESS, StringComparison.InvariantCultureIgnoreCase)
                                                                            && _.SelectedValues.Any());

            var filterCategoria = Context.ApplicableFilters.FirstOrDefault(_ => _.Table == inputDataFilters_Table
                                                                            && _.FieldName.Equals(Values.HEADER_CATEGORIA, StringComparison.InvariantCultureIgnoreCase)
                                                                            && _.SelectedValues.Any());
            #endregion


            #region Leggo tutte le righe della sorgente a partire da quella immediatamente successiva alla riga con gli headers
            var righeLette = new List<object[]>();
            var currentBusiness = "";
            var lastRow = sourceWorksheet.Dimension.End.Row;
            for (var rowSourceIndex = souceHeadersRow + 1; rowSourceIndex <= lastRow; rowSourceIndex++)
            {
                #region Lettura campo "Business"
                // Leggo il valore sulla colonna "Business"
                var valoreCellaBusiness = sourceWorksheet.Cells[rowSourceIndex, sourceHeadersFirstColumn].Value;

                // Se questa cella è diversa na null, inizia un nuovo gruppo di righe di questo "Business"
                if (valoreCellaBusiness != null)
                {
                    currentBusiness = valoreCellaBusiness.ToString().Trim();
                    continue;
                }

                // applico gli eventuali alias
                currentBusiness = Context.ApplicaAliasToValue(Values.HEADER_BUSINESS, currentBusiness);

                // Applico il filtro: se il valore non è presente tra i valori selezionati, la riga viene saltata
                if (filterBusiness != null && !filterBusiness.SelectedValues.Any(_ => _.Equals(currentBusiness, StringComparison.InvariantCultureIgnoreCase)))
                { continue; }
                #endregion


                #region Lettura campo "Categoria"
                var categoria = sourceWorksheet.Cells[rowSourceIndex, sourceHeadersFirstColumn + 1].Value.ToString()
                            ?? throw new Exception("Column 'Categoria' cannot be empty");

                // applico gli eventuali alias
                categoria = Context.ApplicaAliasToValue(Values.HEADER_CATEGORIA, categoria);

                // Applico il filtro: se il valore non è presente tra i valori selezionati, la riga viene saltata
                if (filterCategoria != null && !filterCategoria.SelectedValues.Any(_ => _.Equals(categoria, StringComparison.InvariantCultureIgnoreCase)))
                { continue; }
                #endregion


                var valoriColonne = new object[9];
                valoriColonne[0] = currentBusiness;
                valoriColonne[1] = categoria;
                #region Lettura delle 7 colonne numeriche
                for (var col = 3; col <= 9; col++)
                {
                    var value = sourceWorksheet.Cells[rowSourceIndex, sourceHeadersFirstColumn + col - 1].Value;

                    // Sostituisco i null con 0
                    if (value == null)
                    { value = (double)0; }

                    var doubleValue = value as double?;
                    if (!doubleValue.HasValue)
                    {
                        throw new Exception("Cella con valore non decimal");
                    }

                    valoriColonne[col - 1] = doubleValue.Value;
                }
                #endregion

                righeLette.Add(valoriColonne);
            }
            #endregion



            // Rappresenta la riga del foglio di destinazione in cui scrivere la prossima riga
            var destRowIndex = destHeadersRow + 1;

            #region Allungo la tabella del numero di righe necessarie
            var numeroRigheDaAggiungere = righeLette.Count;
            // Per l'aggiunta delle righe parto sempre dalla prima immediatamente dopo gli headers per asicurarmi di preservare le formule inserendo nuove righe            
            destWorksheet.InsertRow(destRowIndex, numeroRigheDaAggiungere);
            infoRowsDestinazione.Aggiunte = numeroRigheDaAggiungere;
            #endregion


            #region Scrivo i dati importati nelle righe appena create
            destWorksheet.Cells[destRowIndex, destHeadersFirstColumn].LoadFromArrays(righeLette);
            destRowIndex += numeroRigheDaAggiungere;
            #endregion


            #region Cancellazione delle righe in più, ovvero quelle gia esistenti che sono shiftate verso il basso

            var numeroRigheDaCancellare = destWorksheet.Dimension.Rows - destHeadersRow - numeroRigheDaAggiungere;
            if (numeroRigheDaCancellare > 0)
            { destWorksheet.DeleteRow(destRowIndex, numeroRigheDaCancellare, true); }
            infoRowsDestinazione.Eliminate = numeroRigheDaCancellare;
            #endregion


            if (numeroRigheDaAggiungere == 0)
            {
                throw new ManagedException(
                    filePath: sourceFileEPPlusHelper.FilePathInUse,
                    fileType: sourceFileType,
                    //
                    worksheetName: sourceWorksheetName,
                    cellRow: null,
                    cellColumn: null,
                    valueHeader: ValueHeaders.None,
                    value: null,
                    //
                    errorType: ErrorTypes.NoDataAvailable,
                    userMessage: string.Format(UserErrorMessages.NoDataAvailableFromFileAfterFilters, sourceFileType)
                    );
            }

            // Log delle informazioni
            infoRowsDestinazione.Preservate = 0;
            infoRowsDestinazione.Riutilizzate = 0;
            infoRowsDestinazione.Finali = destWorksheet.Dimension.End.Row - destHeadersRow;
            infoRowsDestinazione.VerificaCoerenzaValori();
            Context.DebugInfoLogger.LogRigheSourceFiles(sourceFileType, infoRowsDestinazione);

            //destWorksheet.Select(destWorksheet.Cells[1, 1]);
            sourceFileEPPlusHelper.Close();
        }
    }
}