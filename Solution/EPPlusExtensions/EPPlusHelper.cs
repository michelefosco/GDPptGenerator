using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace EPPlusExtensions
{
    public class EPPlusHelper
    {
        private ExcelPackage _excelPackage;
        private readonly Color _backgroundColorCellaInEvidenza;
        private readonly Color _backgroundColorCellaInErrore;


        public string FilePathInUse { get; private set; }

        public EPPlusHelper()
        {
            _backgroundColorCellaInEvidenza = System.Drawing.Color.Yellow;
            _backgroundColorCellaInErrore = System.Drawing.Color.Red;
        }


        public void AddNewHeaderRow(string worksheetName, object value_01, object value_02 = null, object value_03 = null, object value_04 = null, object value_05 = null, object value_06 = null, object value_07 = null, object value_08 = null, object value_09 = null, object value_10 = null, object value_11 = null, object value_12 = null, object value_13 = null, object value_14 = null, object value_15 = null)
        {
            AddNewRowWithContent(worksheetName, true, value_01, value_02, value_03, value_04, value_05, value_06, value_07, value_08, value_09, value_10, value_11, value_12, value_13, value_14, value_15);
        }

        public void AddNewContentRow(string worksheetName, object value_01, object value_02 = null, object value_03 = null, object value_04 = null, object value_05 = null, object value_06 = null, object value_07 = null, object value_08 = null, object value_09 = null, object value_10 = null, object value_11 = null, object value_12 = null, object value_13 = null, object value_14 = null, object value_15 = null)
        {
            AddNewRowWithContent(worksheetName, false, value_01, value_02, value_03, value_04, value_05, value_06, value_07, value_08, value_09, value_10, value_11, value_12, value_13, value_14, value_15);
        }

        private void AddNewRowWithContent(string worksheetName, bool isHeader, object value_01, object value_02 = null, object value_03 = null, object value_04 = null, object value_05 = null, object value_06 = null, object value_07 = null, object value_08 = null, object value_09 = null, object value_10 = null, object value_11 = null, object value_12 = null, object value_13 = null, object value_14 = null, object value_15 = null)
        {
            var currentWorksheet = GetWorksheet(worksheetName, true);
            var rowToBeUsed = (currentWorksheet.Dimension != null)
                            ? currentWorksheet.Dimension.End.Row + 1
                            : 1;

            var lastColumnWithData = 0;
            if (value_01 != null) { currentWorksheet.Cells[rowToBeUsed, 1].Value = value_01; lastColumnWithData = 1; }
            if (value_02 != null) { currentWorksheet.Cells[rowToBeUsed, 2].Value = value_02; lastColumnWithData = 2; }
            if (value_03 != null) { currentWorksheet.Cells[rowToBeUsed, 3].Value = value_03; lastColumnWithData = 3; }
            if (value_04 != null) { currentWorksheet.Cells[rowToBeUsed, 4].Value = value_04; lastColumnWithData = 4; }
            if (value_05 != null) { currentWorksheet.Cells[rowToBeUsed, 5].Value = value_05; lastColumnWithData = 5; }
            if (value_06 != null) { currentWorksheet.Cells[rowToBeUsed, 6].Value = value_06; lastColumnWithData = 6; }
            if (value_07 != null) { currentWorksheet.Cells[rowToBeUsed, 7].Value = value_07; lastColumnWithData = 7; }
            if (value_08 != null) { currentWorksheet.Cells[rowToBeUsed, 8].Value = value_08; lastColumnWithData = 8; }
            if (value_09 != null) { currentWorksheet.Cells[rowToBeUsed, 9].Value = value_09; lastColumnWithData = 9; }
            if (value_10 != null) { currentWorksheet.Cells[rowToBeUsed, 10].Value = value_10; lastColumnWithData = 10; }
            if (value_11 != null) { currentWorksheet.Cells[rowToBeUsed, 11].Value = value_11; lastColumnWithData = 11; }
            if (value_12 != null) { currentWorksheet.Cells[rowToBeUsed, 12].Value = value_12; lastColumnWithData = 12; }
            if (value_13 != null) { currentWorksheet.Cells[rowToBeUsed, 13].Value = value_13; lastColumnWithData = 13; }
            if (value_14 != null) { currentWorksheet.Cells[rowToBeUsed, 14].Value = value_14; lastColumnWithData = 14; }
            if (value_15 != null) { currentWorksheet.Cells[rowToBeUsed, 15].Value = value_15; lastColumnWithData = 15; }

            if (isHeader)
            {
                var excelRange = currentWorksheet.Cells[rowToBeUsed, 1, rowToBeUsed, lastColumnWithData];
                excelRange.Style.Font.Bold = true;
                excelRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
        }

        public void AutoFitColumns(string worksheetName)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            if (currentWorksheet.Dimension != null)
            {
                currentWorksheet.Cells[currentWorksheet.Dimension.Address].AutoFitColumns();
            }
        }

        public void BorderAllContent(string worksheetName)
        {
            var currentWorksheet = GetWorksheet(worksheetName);

            if (currentWorksheet.Dimension != null)
            {
                currentWorksheet.Cells[currentWorksheet.Dimension.Address].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                currentWorksheet.Cells[currentWorksheet.Dimension.Address].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                currentWorksheet.Cells[currentWorksheet.Dimension.Address].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                currentWorksheet.Cells[currentWorksheet.Dimension.Address].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                currentWorksheet.Cells[currentWorksheet.Dimension.Address].Style.Border.BorderAround(ExcelBorderStyle.Thick);
            }
        }

        public void RemoveEntireRow(string worksheetName, int row)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            //currentWorksheet.DeleteRow(row,1,true);
            currentWorksheet.DeleteRow(row);
        }

        public void SetValuesOnRow(string worksheetName, int row, int firstColumnToBeUsed, object value_01, object value_02 = null, object value_03 = null, object value_04 = null, object value_05 = null, object value_06 = null, object value_07 = null, object value_08 = null, object value_09 = null, object value_10 = null)
        {
            var currentWorksheet = GetWorksheet(worksheetName);

            if (value_01 != null) { currentWorksheet.Cells[row, firstColumnToBeUsed + 0].Value = value_01; }
            if (value_02 != null) { currentWorksheet.Cells[row, firstColumnToBeUsed + 1].Value = value_02; }
            if (value_03 != null) { currentWorksheet.Cells[row, firstColumnToBeUsed + 2].Value = value_03; }
            if (value_04 != null) { currentWorksheet.Cells[row, firstColumnToBeUsed + 3].Value = value_04; }
            if (value_05 != null) { currentWorksheet.Cells[row, firstColumnToBeUsed + 4].Value = value_05; }
            if (value_06 != null) { currentWorksheet.Cells[row, firstColumnToBeUsed + 5].Value = value_06; }
            if (value_07 != null) { currentWorksheet.Cells[row, firstColumnToBeUsed + 6].Value = value_07; }
            if (value_08 != null) { currentWorksheet.Cells[row, firstColumnToBeUsed + 7].Value = value_08; }
            if (value_09 != null) { currentWorksheet.Cells[row, firstColumnToBeUsed + 8].Value = value_09; }
            if (value_10 != null) { currentWorksheet.Cells[row, firstColumnToBeUsed + 9].Value = value_10; }
        }

        public void CleanCellsContent(string worksheetName, int fromRow, int fromCol, int toRow, int toCol)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            //var excelRange = currentWorksheet.Cells[fromRow, fromCol, toRow, toCol];
            // excelRange.Clear();
            for (var row = fromRow; row <= toRow; row++)
            {
                for (var col = fromCol; col <= toCol; col++)
                {
                    currentWorksheet.Cells[row, col].Value = null;
                }
            }
        }

        public bool ExistsGetWorksheet(string worksheetName)
        {
            var worksheet = _excelPackage.Workbook.Worksheets[worksheetName];
            return !(worksheet == null);
        }

        private ExcelWorksheet GetWorksheet(string worksheetName, bool addIfMissing = false)
        {
            var worksheet = _excelPackage.Workbook.Worksheets[worksheetName];
            if (worksheet == null)
            {
                if (addIfMissing)
                {
                    worksheet = _excelPackage.Workbook.Worksheets.Add(worksheetName);
                }
                else
                {
                    throw new Exception($"Worksheet '{worksheetName}' is missing");
                }
            }
            return worksheet;
        }

        public int GetRowsLimit(string worksheetName)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            return currentWorksheet.Dimension.End.Row;
        }

        public int GetLastUsedRowForColumn(string worksheetName, int rowToStartSearchFrom, int colToBeChecked, bool ignoreBlanks = true)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            var lastUsuedRow = 0;

            var rigaCorrente = rowToStartSearchFrom;
            // Scorro tutta la colonna e memorizzo la posizione dell'ultima cella con un valore o una formula
            while (rigaCorrente <= currentWorksheet.Dimension.End.Row)
            {
                var hasFormula = currentWorksheet.Cells[rigaCorrente, colToBeChecked].Formula != "";
                var hasValue = ignoreBlanks
                            ? currentWorksheet.Cells[rigaCorrente, colToBeChecked].Value != null && !string.IsNullOrWhiteSpace(currentWorksheet.Cells[rigaCorrente, colToBeChecked].Value.ToString().Trim())
                            : currentWorksheet.Cells[rigaCorrente, colToBeChecked].Value != null;

                if (hasFormula || hasValue)
                {
                    // aggiorlo l'ultima riga in cui ho trovato una cella (nella colonna controllata) con un valore o una formula
                    lastUsuedRow = rigaCorrente;
                }
                rigaCorrente++;
            }

            return lastUsuedRow;
        }

        public int GetFirstRowWithSpecificValue(string worksheetName, int rowToStartSearchFrom, int colToBeChecked, string valueToFind)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            var rowWithTheValueFound = -1;

            var rigaCorrente = rowToStartSearchFrom;
            // Scorro tutta la colonna e memorizzo la posizione dell'ultima cella con il valore cercato
            while (rigaCorrente <= currentWorksheet.Dimension.End.Row)
            {
                if (currentWorksheet.Cells[rigaCorrente, colToBeChecked].Value != null)
                {
                    if (currentWorksheet.Cells[rigaCorrente, colToBeChecked].Value.ToString().Trim().Equals(valueToFind, StringComparison.OrdinalIgnoreCase))
                    {
                        rowWithTheValueFound = rigaCorrente;
                    }
                }
                rigaCorrente++;
            }

            return rowWithTheValueFound;
        }

        public int GetFirstRowWithAnyValue(string worksheetName, int rowToStartSearchFrom, int colToBeChecked, bool ignoreBlanks = true, string ignoreThisText = null)
        {
            var currentWorksheet = GetWorksheet(worksheetName);

            var rigaCorrente = rowToStartSearchFrom;
            while (rigaCorrente <= currentWorksheet.Dimension.End.Row)
            {
                if (currentWorksheet.Cells[rigaCorrente, colToBeChecked].Value != null)
                {
                    //ho trovato una cella con valore

                    // verifico le opzioni da ignorare: opzione "Ignora i blank"
                    if (ignoreBlanks && string.IsNullOrWhiteSpace(currentWorksheet.Cells[rigaCorrente, colToBeChecked].Value.ToString().Trim()))
                    {
                        rigaCorrente++;
                        continue; // ignoro i black
                    }

                    // verifico le opzioni da ignorare: opzione "Ignora un determinato testo"
                    if (ignoreThisText != null && currentWorksheet.Cells[rigaCorrente, colToBeChecked].Value.ToString().Equals(ignoreThisText, StringComparison.OrdinalIgnoreCase))
                    {
                        rigaCorrente++;
                        continue; // ignoro le celle contententi il testo della variabile ignoreThisText
                    }

                    return rigaCorrente;
                }

                rigaCorrente++;
            }

            // ho raggiunto il limite del foglio senza aver indivudiato una cella con un valore
            return -1;
        }

        public int GetFirstEmptyRow(string worksheetName, int rowToStartSearchFrom, int colToBeChecked)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            var rigaCorrente = rowToStartSearchFrom;

            // Scorro le righe finchè sono diverse da Null. Mi fermo quando trovo un valore o una formula
            while (currentWorksheet.Cells[rigaCorrente, colToBeChecked].Value != null || currentWorksheet.Cells[rigaCorrente, colToBeChecked].Formula != "")
            { rigaCorrente++; }

            return rigaCorrente;
        }

        public int GetFirstEmptyRow(string worksheetName, int rowToStartSearchFrom, int colToBeChecked1, int colToBeChecked2, int colToBeChecked3)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            var rigaCorrente = rowToStartSearchFrom;

            // Scorro le righe finchè sono diverse da Null. Mi fermo quando trovo un valore o una formula
            while ((currentWorksheet.Cells[rigaCorrente, colToBeChecked1].Value != null || currentWorksheet.Cells[rigaCorrente, colToBeChecked1].Formula != "")
                && (currentWorksheet.Cells[rigaCorrente, colToBeChecked2].Value != null || currentWorksheet.Cells[rigaCorrente, colToBeChecked2].Formula != "")
                && (currentWorksheet.Cells[rigaCorrente, colToBeChecked3].Value != null || currentWorksheet.Cells[rigaCorrente, colToBeChecked3].Formula != ""))
            { rigaCorrente++; }

            return rigaCorrente;
        }

        public object GetValue(string worksheetName, int row, int col)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            return currentWorksheet.Cells[row, col].Value;
        }

        public List<object> GetValuesFromColumn(string worksheetName, int col, int rowFrom, int rowTo)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            var values = new List<object>();
            for (int rigaCorrente = rowFrom; rigaCorrente <= rowTo; rigaCorrente++)
            {
                values.Add(currentWorksheet.Cells[rigaCorrente, col].Value);
            }
            return values;
        }

        /// <summary>
        /// Legge i valori presenti nella riga (utile per le intestazioni)
        /// </summary>
        /// <param name="worksheetName"></param>
        /// <param name="row"></param>
        /// <param name="colFrom"></param>
        /// <returns></returns>
        public List<string> GetHeaders(string worksheetName, int row, int colFrom = 1)
        {
            var currentWorksheet = GetWorksheet(worksheetName);

            var values = new List<string>();
            for (int colonnaCorrente = colFrom; colonnaCorrente <= currentWorksheet.Dimension.Columns; colonnaCorrente++)
            {
                var value = currentWorksheet.Cells[row, colonnaCorrente].Value.ToString();
                if (!string.IsNullOrEmpty(value))
                { values.Add(value); }
            }

            return values;
        }

        public List<string> GetValuesFromColumnsWithHeader(string worksheetName, int headersRow, string headerValue, int colFrom = 1)
        {
            var currentWorksheet = GetWorksheet(worksheetName);


            for (int colonnaCorrente = colFrom; colonnaCorrente <= currentWorksheet.Dimension.Columns; colonnaCorrente++)
            {
                var value = currentWorksheet.Cells[headersRow, colonnaCorrente].Value.ToString();
                if (value.Equals(headerValue, StringComparison.CurrentCultureIgnoreCase))
                {
                    // ho trovato la colonna che mi interessa, leggo i valori
                    var values = new List<string>();
                    for (int rigaCorrente = headersRow + 1; rigaCorrente <= currentWorksheet.Dimension.End.Row; rigaCorrente++)
                    {
                        var cellValue = currentWorksheet.Cells[rigaCorrente, colonnaCorrente].Value;
                        if (cellValue != null)
                        { values.Add(cellValue.ToString()); }
                    }
                    return values;
                }
            }

            //todo: se arrivo qui non ho trovato l'header che andava prima controllato
            throw new Exception($"Header '{headerValue}' not found in worksheet '{worksheetName}'");
        }

        public string GetFormula(string worksheetName, int row, int col)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            return currentWorksheet.Cells[row, col].Formula;
        }

        public bool IsFormula(string worksheetName, int row, int col)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            return !string.IsNullOrEmpty(currentWorksheet.Cells[row, col].Formula);
        }

        public bool WorksheetExists(string worksheetName)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            return !(currentWorksheet.Dimension == null);
        }

        public string GetString(string worksheetName, int row, int col, bool trimString = true)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            var value = currentWorksheet.Cells[row, col].Value;
            if (value == null) { return null; }

            return trimString ? value.ToString().Trim()
                              : value.ToString();
        }

        public List<string> GetWorksheetNames()
        {
            var list = _excelPackage.Workbook.Worksheets.Select(_ => _.Name.ToUpper()).ToList();
            return _excelPackage.Workbook.Worksheets.Select(_ => _.Name.ToUpper()).ToList();
        }

        public void RemoveWorksheet(string worksheetName)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            _excelPackage.Workbook.Worksheets.Delete(currentWorksheet);
        }

        public bool Create(string filePath, string firstWorkSheetName)
        {
            try
            {
                _excelPackage = new ExcelPackage(new FileInfo(filePath));
                _excelPackage.Workbook.Worksheets.Add(firstWorkSheetName);
                _excelPackage.Save();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool Open(string filePath)
        {
            // resetto il nome del file aperto
            FilePathInUse = null;

            if (!File.Exists(filePath))
            { return false; }

            try
            {
                _excelPackage = new ExcelPackage(new FileInfo(filePath));
                // file aperto correttamente
                FilePathInUse = filePath;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public int SetRefreshOnLoadForAllPivotTables()
        {
            var updatedPivotTables = 0;
            foreach (var currentWorksheet in _excelPackage.Workbook.Worksheets)
            {
                var pivotTables = currentWorksheet.PivotTables;
                if (pivotTables != null && pivotTables.Any())
                {
                    foreach (var pivotTable in pivotTables)
                    {
                        var currentValue = pivotTable.CacheDefinition.CacheDefinitionXml.DocumentElement?.GetAttribute("refreshOnLoad");
                        if (string.IsNullOrEmpty(currentValue) || currentValue != "1")
                        {
                            pivotTable.CacheDefinition.CacheDefinitionXml.DocumentElement?.SetAttribute("refreshOnLoad", "1");
                            updatedPivotTables++;
                        }
                    }
                }
            }
            return updatedPivotTables;
        }

        public bool Save()
        {
            try
            {
                _excelPackage.Save();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool SaveAs(string newFilePath)
        {
            try
            {
                _excelPackage.SaveAs(new FileInfo(newFilePath));
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public void SetValue(string worksheetName, int row, int col, object value, bool highlight = false)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            currentWorksheet.Cells[row, col].Value = value;
            if (highlight)
            { SetBackgroundColor(worksheetName, row, col, _backgroundColorCellaInEvidenza, 0); }
        }

        public void SetValue(string worksheetName, string address, object value)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            currentWorksheet.Cells[address].Value = value;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="worksheetName"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="backgroundColor"></param>
        /// <param name="tint">Tint a 0 = colore pieno</param>
        public void SetBackgroundColor(string worksheetName, int row, int col, Color backgroundColor, decimal? tint = null)
        {
            var currentWorksheet = GetWorksheet(worksheetName);

            currentWorksheet.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;

            if (tint.HasValue)
            { currentWorksheet.Cells[row, col].Style.Fill.BackgroundColor.Tint = tint.Value; }

            currentWorksheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(backgroundColor);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="worksheetName"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="backgroundColor"></param>
        /// <param name="tint">Tint a 0 = colore pieno</param>
        public void SetBackgroundColor(string worksheetName, int fromRow, int fromCol, int toRow, int toCol, Color backgroundColor, decimal? tint = null)
        {
            var currentWorksheet = GetWorksheet(worksheetName);

            currentWorksheet.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;

            if (tint.HasValue)
            { currentWorksheet.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.Tint = tint.Value; }

            currentWorksheet.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(backgroundColor);
        }


        public void SetThinBorderOnRange(string worksheetName, int fromRow, int fromCol, int toRow, int toCol)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            var excelRange = currentWorksheet.Cells[fromRow, fromCol, toRow, toCol];
            excelRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        }

        public void SetThickRightBorder(string worksheetName, int row, int col)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            currentWorksheet.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thick;
        }

        public void SelectWorksheet(string worksheetName)
        {
            SelectWorksheet(worksheetName, null, null);
        }

        public void SelectWorksheet(string worksheetName, int? row, int? col)
        {
            var currentWorksheet = GetWorksheet(worksheetName);

            if (row.HasValue && col.HasValue)
            {
                var range = currentWorksheet.Cells[row.Value, col.Value];
                currentWorksheet.Select(range);
            }
            else
            {
                currentWorksheet.Select();
            }
        }

        public void SetErrorCell(string worksheetName, int? row, int? col)
        {
            SelectWorksheet(worksheetName, row, col);
            if (row.HasValue && col.HasValue)
            {
                SetBackgroundColor(worksheetName, row.Value, col.Value, _backgroundColorCellaInErrore, 0);
            }
        }

        public void SetFormula(string worksheetName, int row, int col, string formula)
        {
            var currentWorksheet = GetWorksheet(worksheetName);

            currentWorksheet.Cells[row, col].Formula = formula;
        }

        public void ConcatenaTestoFormula(string worksheetName, int row, int col, string testoDaConcatenareAllaFormula)
        {
            var currentWorksheet = GetWorksheet(worksheetName);

            var currentFormula = currentWorksheet.Cells[row, col].Formula;
            var newFormula = currentFormula + testoDaConcatenareAllaFormula;
            currentWorksheet.Cells[row, col].Formula = newFormula;
        }

        public void CopiaFormula(string worksheetName, int originRow, int originCol, int destinationRow, int destinationCol)
        {
            var currentWorksheet = GetWorksheet(worksheetName);

            // Copia il contenuto da FROM a TO
            var originCell = currentWorksheet.Cells[originRow, originCol];
            var destinationCell = currentWorksheet.Cells[destinationRow, destinationCol];

            // preserva lo stile della cella di destinazione
            var originalStyleID = destinationCell.StyleID;

            originCell.Copy(destinationCell);

            // riapplica lo stile originale alla cella di destinazione
            destinationCell.StyleID = originalStyleID;
        }

        public void OrdinaTabella(string worksheetName, int fromRow, int fromCol, int toRow, int toCol, int sortColumnIndexZeroBased = 0)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            var excelRange = currentWorksheet.Cells[fromRow, fromCol, toRow, toCol];
            //            excelRange.Sort(sortColumnIndexZeroBased);
            excelRange.Sort(new int[1] { sortColumnIndexZeroBased }, null, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.CompareOptions.StringSort);
        }

        public void AggiungiRighe(string worksheetName, int fromRow, int numberOfRows = 1)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            currentWorksheet.InsertRow(fromRow, numberOfRows);
        }

        public void NascondiRighe(string worksheetName, int fromRow, int toRow, bool hidden)
        {
            var currentWorksheet = GetWorksheet(worksheetName);
            for (int rowNumber = fromRow; rowNumber <= toRow; rowNumber++)
            {
                currentWorksheet.Row(rowNumber).Hidden = hidden;
            }
        }


        public void Close()
        {
            _excelPackage.Dispose();
        }


    }
}