using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FilesEditor.Steps.BuildPresentation
{
    /// <summary>
    /// 
    /// </summary>
    internal class Step_ImportaDati_BudgetAndForecast : StepBase
    {
        public override string StepName => "Step_ImportaDati_BudgetAndForecast";

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
        public Step_ImportaDati_BudgetAndForecast(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificStepTask()
        {
            ImportaSourceFile(
                    sourceFileType: FileTypes.Budget,
                    sourceFilePath: Context.FileBudgetPath,
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
                    sourceFilePath: Context.FileForecastPath,
                    // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome
                    sourceWorksheetName: null, // WorksheetNames.SOURCEFILE_FORECAST_DATA,
                    souceHeadersRow: Context.Configurazione.SOURCE_FILES_FORECAST_HEADERS_ROW,
                    sourceHeadersFirstColumn: Context.Configurazione.SOURCE_FILES_FORECAST_HEADERS_FIRST_COL,
                    //
                    destWorksheetName: WorksheetNames.DATASOURCE_FORECAST_DATA,
                    destHeadersRow: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_ROW,
                    destHeadersFirstColumn: Context.Configurazione.DATASOURCE_FORECAST_HEADERS_FIRST_COL
                );

            return EsitiFinali.Undefined; // Step intermedio, non ritorna alcun esito
        }

        private void ImportaSourceFile(
                FileTypes sourceFileType,
                string sourceFilePath,
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

            // variabili per il conteggio delle righe Eliminate, Aggiunge e Preservate (quando in modalità append su Superdettagli)
            int totRighePreservate = 0;
            int totRigheEliminate = 0;
            int totRigheAggiunte = 0;

            #region WorkSheets sorgente e destinazione
            // Foglio sorgente
            var packageSource = new ExcelPackage(new FileInfo(sourceFilePath));

            // 06/11/2025, Francesco chiede di usare sempre il 1° foglio presente nel file, indipendentemente dal nome
            //var sourceWorksheet = packageSource.Workbook.Worksheets[sourceWorksheetName];
            var sourceWorksheet = packageSource.Workbook.Worksheets[1];

            // Foglio destinazione
            var worksheetDest = Context.DataSourceEPPlusHelper.ExcelPackage.Workbook.Worksheets[destWorksheetName];
            #endregion

            #region Individuo gli evemtuali filtri da applicare
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


            #region Scorro tutte le righe della sorgente a partire da quella immediatamente successiva alla riga con gli headers
            var righe = new List<RigaBudgetForecast>();
            var currentBusiness = "";
            for (var rowSourceIndex = souceHeadersRow + 1; rowSourceIndex <= sourceWorksheet.Dimension.End.Row; rowSourceIndex++)
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


                #region Lettura delle 7 colonne numeriche
                var columns = new double[7];
                for (var col = 1; col <= 7; col++)
                {
                    var value = sourceWorksheet.Cells[rowSourceIndex, sourceHeadersFirstColumn + 1 + col].Value;

                    // Sostituisco i null con 0
                    if (value == null)
                    { value = (double)0; }

                    var doubleValue = value as double?;
                    if (!doubleValue.HasValue)
                    {
                        throw new Exception("Cella con valore non decimal");
                    }

                    columns[col - 1] = doubleValue.Value;
                }
                #endregion


                var riga = new RigaBudgetForecast(currentBusiness, categoria, columns);
                righe.Add(riga);
            }
            #endregion


            // Per l'aggiunta delle righe parto sempre dalla prima immediatamente dopo gli headers per asicurarmi di preservare le formule inserendo nuove righe
            // Rappresenta la riga del foglio di destinazione in cui scrivere la prossima riga
            var destRowIndex = destHeadersRow + 1;
            foreach (var riga in righe)
            {
                #region Allungo la tabella di un riga in modo da conservare le formule
                worksheetDest.InsertRow(destRowIndex, 1);
                totRigheAggiunte++;
                #endregion

                worksheetDest.Cells[destRowIndex, destHeadersFirstColumn].Value = riga.Business;
                worksheetDest.Cells[destRowIndex, destHeadersFirstColumn + 1].Value = riga.Categoria;
                for (int col = 1; col <= 7; col++)
                { worksheetDest.Cells[destRowIndex, destHeadersFirstColumn + 1 + col].Value = riga.Columns[col - 1]; }
                destRowIndex++;
            }

            // ritorno indietro di uno per mantenere il significato del valore (Rappresenta la riga del foglio di destinazione in cui scrivere la prossima riga) 
            destRowIndex--;


            #region Cancellazione delle righe in più, ovvero quelle gia esistenti ma non più utilizzate (esempio ho meno righe dell'aggiornamento precedente)
            // aggiungo al conteggio le righe preservate, in modo da non canellarle
            destRowIndex += totRighePreservate;

            // la cancellazione deve avvenire dall'ultima riga indietro in quanto le righe eliminate shiftano verso il basso e gli indici delle righe vengono aggiornati
            for (var rowIndex = worksheetDest.Dimension.Rows; rowIndex > destRowIndex; rowIndex--)
            {
                worksheetDest.DeleteRow(rowIndex, 1, true);
                totRigheEliminate++;
            }
            #endregion

            if (totRighePreservate + totRigheAggiunte == 0)
            {
                throw new ManagedException(
                    filePath: sourceFilePath,
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
            Context.DebugInfoLogger.LogRigheSourceFiles(sourceFileType, totRighePreservate, totRigheEliminate, totRigheAggiunte);
        }
    }
}