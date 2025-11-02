using FilesEditor.Constants;
using FilesEditor.Entities;
using FilesEditor.Entities.Exceptions;
using FilesEditor.Enums;
using FilesEditor.Helpers;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FilesEditor.Steps.BuildPresentation
{
    internal class Step_ImportaDatiDa_BudgetAndForecast : StepBase
    {
        public Step_ImportaDatiDa_BudgetAndForecast(StepContext context) : base(context)
        { }

        internal override EsitiFinali DoSpecificTask()
        {
            Context.DebugInfoLogger.LogStepContext("Step_ImportaDatiDaSourceFiles", Context);

            ImportaSourceFile(
                    sourceFileType: FileTypes.Budget,
                    sourceFilePath: Context.FileBudgetPath,
                    sourceWorksheetName: WorksheetNames.SOURCEFILE_BUDGET_DATA,
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
                    sourceWorksheetName: WorksheetNames.SOURCEFILE_FORECAST_DATA,
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
            // variabili per il conteggio delle righe Eliminate, Aggiunge e Preservate (quando in modalità append su Superdettagli)
            int totRighePreservate = 0;
            int totRigheEliminate = 0;
            int totRigheAggiunte = 0;

            #region WorkSheets sorgente e destinazione
            // Foglio sorgente
            var packageSource = new ExcelPackage(new FileInfo(sourceFilePath));
            var worksheetSource = packageSource.Workbook.Worksheets[sourceWorksheetName];

            // Foglio destinazione
            var worksheetDest = Context.EpplusHelperDataSource.ExcelPackage.Workbook.Worksheets[destWorksheetName];
            #endregion

            #region Preparo una struttura più snella che contenga le informazioni su filtri
            InputDataFilters_Tables inputDataFilters_Table;
            switch (sourceFileType)
            {
                case FileTypes.Budget:
                    inputDataFilters_Table = InputDataFilters_Tables.BUDGET;
                    break;
                case FileTypes.Forecast:
                    inputDataFilters_Table = InputDataFilters_Tables.FORECAST;
                    break;
                //case FileTypes.RunRate:
                //case FileTypes.SuperDettagli:
                default:
                    throw new ArgumentOutOfRangeException();
            }
            var filtersWithSelectedValues = Context.ApplicableFilters.Where(af => af.Table == inputDataFilters_Table && af.SelectedValues.Any()).ToList();
            var filterBusiness = filtersWithSelectedValues.FirstOrDefault(_ => _.FieldName.Equals(Values.HEADER_BUSINESS, StringComparison.InvariantCultureIgnoreCase));
            var filterCategoria = filtersWithSelectedValues.FirstOrDefault(_ => _.FieldName.Equals(Values.HEADER_CATEGORIA, StringComparison.InvariantCultureIgnoreCase));
            #endregion

            
            #region Scorro tutte le righe della sorgente a partire da quella immediatamente successiva alla riga con gli headers
            var righe = new List<RigaBudgetForecast>();
            var currentBusiness = "";
            for (var rowSourceIndex = souceHeadersRow + 1; rowSourceIndex <= worksheetSource.Dimension.End.Row; rowSourceIndex++)
            {
                #region Lettura campo "Business"
                // Leggo il valore sulla colonna "Business"
                var valoreCellaBusiness = worksheetSource.Cells[rowSourceIndex, sourceHeadersFirstColumn].Value;

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
                var categoria = worksheetSource.Cells[rowSourceIndex, sourceHeadersFirstColumn + 1].Value.ToString();
                if (categoria == null)
                {
                    //todo: sollevare eccezione
                    throw new Exception("Categoria non può essere vuoto");
                }

                // applico gli eventuali alias
                categoria = Context.ApplicaAliasToValue(Values.HEADER_CATEGORIA, categoria);

                // Applico il filtro: se il valore non è presente tra i valori selezionati, la riga viene saltata
                if (filterCategoria != null && !filterCategoria.SelectedValues.Any(_ => _.Equals(categoria, StringComparison.InvariantCultureIgnoreCase)))
                { continue; }

                #endregion


                #region Lettura delle 7 colonne numeriche
                double[] columns = new double[7];
                for (int col = 1; col <= 7; col++)
                {
                    var value = worksheetSource.Cells[rowSourceIndex, sourceHeadersFirstColumn + 1 + col].Value;

                    // Sostituisco i null con 0
                    if (value == null)
                    { value = (double)0; }

                    double? doubleValue = value as double?;
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
            for (int rowIndex = worksheetDest.Dimension.Rows; rowIndex > destRowIndex; rowIndex--)
            {
                worksheetDest.DeleteRow(rowIndex, 1, true);
                totRigheEliminate++;
            }
            #endregion

            // Log delle informazioni
            Context.DebugInfoLogger.LogRigheSourceFiles(sourceFileType, totRighePreservate, totRigheEliminate, totRigheAggiunte);
        }
    }
}