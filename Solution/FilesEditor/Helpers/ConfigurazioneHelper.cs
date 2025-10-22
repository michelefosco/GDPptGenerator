using FilesEditor.Entities;
using FilesEditor.Enums;

namespace FilesEditor.Helpers
{
    public class ConfigurazioneHelper
    {
        public static Configurazione GetConfigurazioneDefault()
        {
            var configurazione = new Configurazione();

            #region File DataSource
            // Foglio configurazione - Filtri
            configurazione.DATASOURCE_CONFIG_FILTERS_FIRST_DATA_ROW = 4;
            configurazione.DATASOURCE_CONFIG_FILTERS_TABLE_COL = (int)ColumnIDS.K;
            configurazione.DATASOURCE_CONFIG_FILTERS_FIELD_COL = (int)ColumnIDS.L;

            // Foglio configurazione - Slide da generare
            configurazione.DATASOURCE_CONFIG_SLIDES_FIRST_DATA_ROW = 4;
            configurazione.DATASOURCE_CONFIG_SLIDES_POWERPOINTFILE_COL = (int)ColumnIDS.A;
            configurazione.DATASOURCE_CONFIG_SLIDES_TITLE_COL = (int)ColumnIDS.B;
            configurazione.DATASOURCE_CONFIG_SLIDES_CONTENT_1_COL = (int)ColumnIDS.C;
            configurazione.DATASOURCE_CONFIG_SLIDES_CONTENT_2_COL = (int)ColumnIDS.D;
            configurazione.DATASOURCE_CONFIG_SLIDES_CONTENT_3_COL = (int)ColumnIDS.E;
            configurazione.DATASOURCE_CONFIG_SLIDES_LAYOUT_COL = (int)ColumnIDS.F;

            // Fogli Printables - PrinteAreas
            configurazione.DATASOURCE_PRINTABLE_ITEMS_PRINT_AREA_ROW = 2;
            configurazione.DATASOURCE_PRINTABLE_ITEMS_PRINT_AREA_COL = (int)ColumnIDS.B;

            // Fogli "Alias" ("Alias-Business TMP", "Alias-Categoria"
            configurazione.DATASOURCE_ALIAS_WORKSHEETS_FIRST_DATA_ROW = 3;
            configurazione.DATASOURCE_ALIAS_WORKSHEETS_RAW_VALUES_COL = 1;
            configurazione.DATASOURCE_ALIAS_WORKSHEETS_NEW_VALUES_COL = 2;

            // Foglio Budget
            configurazione.DATASOURCE_BUDGET_HEADERS_ROW = 2;
            configurazione.DATASOURCE_BUDGET_HEADERS_FIRST_COL = (int)ColumnIDS.K;

            // Foglio Forecast
            configurazione.DATASOURCE_FORECAST_HEADERS_ROW = 3;
            configurazione.DATASOURCE_FORECAST_HEADERS_FIRST_COL = (int)ColumnIDS.P;

            // Foglio RunRate
            configurazione.DATASOURCE_RUNRATE_HEADERS_ROW = 1;
            configurazione.DATASOURCE_RUNRATE_HEADERS_FIRST_COL = (int)ColumnIDS.A;

            // Foglio Superdettagli
            configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_ROW = 2;
            configurazione.DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL = (int)ColumnIDS.AE;
            #endregion


            #region Input files
            // Input file Budget
            configurazione.INPUT_FILES_BUDGET_HEADERS_ROW = 2;
            configurazione.INPUT_FILES_BUDGET_HEADERS_FIRST_COL = (int)ColumnIDS.A;

            // Input file Forecast
            configurazione.INPUT_FILES_FORECAST_HEADERS_ROW = 3;
            configurazione.INPUT_FILES_FORECAST_HEADERS_FIRST_COL = (int)ColumnIDS.A;

            // Input file RunRate
            configurazione.INPUT_FILES_RUNRATE_HEADERS_ROW = 1;
            configurazione.INPUT_FILES_RUNRATE_HEADERS_FIRST_COL = (int)ColumnIDS.A;

            // Input file Superdettagli
            configurazione.INPUT_FILES_SUPERDETTAGLI_HEADERS_ROW = 1;
            configurazione.INPUT_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL = (int)ColumnIDS.A;
            #endregion


            #region Cartella "Debug"
            configurazione.AutoSaveDebugFile = true;
            #endregion

            return configurazione;
        }
    }
}