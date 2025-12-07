namespace FilesEditor.Entities
{
    public class Configurazione
    {
        public bool AutoSaveDebugFile { get; set; }
        public bool ZipBackupFile { get; set; }


        #region DataSource Template

        // Configurazione - Filtri
        public int DATASOURCE_CONFIG_FILTERS_FIRST_DATA_ROW { get; set; }
        public int DATASOURCE_CONFIG_FILTERS_TABLE_COL { get; set; }
        public int DATASOURCE_CONFIG_FILTERS_FIELD_COL { get; set; }

        // Configurazione - Slide da generare
        public int DATASOURCE_CONFIG_SLIDES_FIRST_DATA_ROW { get; set; }
        public int DATASOURCE_CONFIG_SLIDES_POWERPOINTFILE_COL { get; set; }
        public int DATASOURCE_CONFIG_SLIDES_TITLE_COL { get; set; }
        public int DATASOURCE_CONFIG_SLIDES_CONTENT_1_COL { get; set; }
        public int DATASOURCE_CONFIG_SLIDES_CONTENT_2_COL { get; set; }
        public int DATASOURCE_CONFIG_SLIDES_CONTENT_3_COL { get; set; }
        public int DATASOURCE_CONFIG_SLIDES_LAYOUT_COL { get; set; }


        // PrinteAreas
        public int DATASOURCE_PRINTABLE_ITEMS_PRINT_AREA_ROW { get; set; }
        public int DATASOURCE_PRINTABLE_ITEMS_PRINT_AREA_COL { get; set; }


        // Alias-Worksheets definitions
        public int DATASOURCE_ALIAS_WORKSHEETS_FIRST_DATA_ROW { get; set; }
        public int DATASOURCE_ALIAS_WORKSHEETS_RAW_VALUES_COL { get; set; }
        public int DATASOURCE_ALIAS_WORKSHEETS_NEW_VALUES_COL { get; set; }


        #region     Headers
        #region Budgets
        public int DATASOURCE_BUDGET_HEADERS_ROW { get; set; }
        public int DATASOURCE_BUDGET_HEADERS_FIRST_COL { get; set; }

        #endregion

        #region Forecast
        public int DATASOURCE_FORECAST_HEADERS_ROW { get; set; }
        public int DATASOURCE_FORECAST_HEADERS_FIRST_COL { get; set; }

        #endregion

        #region Run rate
        public int DATASOURCE_RUNRATE_HEADERS_ROW { get; set; }
        public int DATASOURCE_RUNRATE_HEADERS_FIRST_COL { get; set; }
        #endregion

        #region Superdettagli
        public int DATASOURCE_SUPERDETTAGLI_HEADERS_ROW { get; set; }
        public int DATASOURCE_SUPERDETTAGLI_HEADERS_FIRST_COL { get; set; }
        public int DATASOURCE_SUPERDETTAGLI_YEAR_COL { get; set; }
        #endregion

        #region Run rate
        public int DATASOURCE_CN43N_HEADERS_ROW { get; set; }
        public int DATASOURCE_CN43N_HEADERS_FIRST_COL { get; set; }
        #endregion

        #endregion

        // Alias-Categoria
        //public int DATASOURCE_ALIAS_CATEGORIA_FIRST_DATA_ROW { get; set; }
        //public int DATASOURCE_ALIAS_CATEGORIA_TMP_RAW_VALUES_COL { get; set; }
        //public int DATASOURCE_ALIAS_CATEGORIA_TMP_NEW_VALUES_COL { get; set; }


        //public int SOURCE_FILES_SUPERDETTAGLI_FIRST_DATA_ROW { get; set; }
        //public int SOURCE_FILES_BUDGET_FIRST_DATA_ROW { get; set; }
        //public int SOURCE_FILES_FORECAST_FIRST_DATA_ROW { get; set; }
        //public int SOURCE_FILES_RUNRATE_FIRST_DATA_ROW { get; set; }
        #endregion

        // ----- INPUT FILES
        #region Budgets
        public int SOURCE_FILES_BUDGET_HEADERS_ROW { get; set; }
        public int SOURCE_FILES_BUDGET_HEADERS_FIRST_COL { get; set; }

        #endregion

        #region Forecast
        public int SOURCE_FILES_FORECAST_HEADERS_ROW { get; set; }
        public int SOURCE_FILES_FORECAST_HEADERS_FIRST_COL { get; set; }

        #endregion

        #region Run rate
        public int SOURCE_FILES_RUNRATE_HEADERS_ROW { get; set; }
        public int SOURCE_FILES_RUNRATE_HEADERS_FIRST_COL { get; set; }
        #endregion

        #region Run rate
        public int SOURCE_FILES_CN43N_HEADERS_ROW { get; set; }
        public int SOURCE_FILES_CN43N_HEADERS_FIRST_COL { get; set; }
        #endregion


        #region Superdettagli
        public int SOURCE_FILES_SUPERDETTAGLI_HEADERS_ROW { get; set; }
        public int SOURCE_FILES_SUPERDETTAGLI_HEADERS_FIRST_COL { get; set; }
        #endregion
    }
}