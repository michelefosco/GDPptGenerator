namespace FilesEditor.Entities
{
    public class Configurazione
    {
        public bool AutoSaveDebugFile { get; set; }

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



   

        //public int INPUT_FILES_SUPERDETTAGLI_FIRST_DATA_ROW { get; set; }
        //public int INPUT_FILES_BUDGET_FIRST_DATA_ROW { get; set; }
        //public int INPUT_FILES_FORECAST_FIRST_DATA_ROW { get; set; }
        //public int INPUT_FILES_RUNRATE_FIRST_DATA_ROW { get; set; }
        #endregion


        #region Superdettagli
        public int INPUT_FILES_SUPERDETTAGLI_HEADERS_ROW { get; set; }
        #endregion

        #region Budgets
        public int INPUT_FILES_BUDGET_HEADERS_ROW { get; set; }

        #endregion

        #region Forecast
        public int INPUT_FILES_FORECAST_HEADERS_ROW { get; set; }

        #endregion

        #region Run rate
        public int INPUT_FILES_RUNRATE_HEADERS_ROW { get; set; }
        #endregion
    }
}