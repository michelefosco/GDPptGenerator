namespace FilesEditor.Entities
{
    public class Configurazione
    {
        public bool AutoSaveDebugFile { get; set; }

        #region DataSource Template

        // "Filtri"
        public int DATASOURCE_TEMPLATE_PPT_CONFIG_FILTERS_FIRST_DATA_ROW { get; set; }
        public int DATASOURCE_TEMPLATE_PPT_CONFIG_FILTERS_TABLE_COL { get; set; }
        public int DATASOURCE_TEMPLATE_PPT_CONFIG_FILTERS_FIELD_COL { get; set; }

        // "Slide da generare
        public int DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_FIRST_DATA_ROW { get; set; }
        public int DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_POWERPOINTFILE_COL { get; set; }
        public int DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_TITLE_COL { get; set; }
        public int DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_CONTENT_1_COL { get; set; }
        public int DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_CONTENT_2_COL { get; set; }
        public int DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_CONTENT_3_COL { get; set; }
        public int DATASOURCE_TEMPLATE_PPT_CONFIG_SLIDES_LAYOUT_COL { get; set; }


        //

        public int DATASOURCE_SUPERDETTAGLI_FIRST_DATA_ROW { get; set; }
        public int DATASOURCE_BUDGET_FIRST_DATA_ROW { get; set; }
        public int DATASOURCE_FORECAST_FIRST_DATA_ROW { get; set; }
        public int DATASOURCE_RANRATE_FIRST_DATA_ROW { get; set; }
        #endregion


        #region Superdettagli
        public int SUPERDETTAGLI_HEADERS_ROW { get; set; }
        #endregion

        #region Budgets
        public int BUDGET_HEADERS_ROW { get; set; }

        #endregion

        #region Forecast
        public int FORECAST_HEADERS_ROW { get; set; }

        #endregion

        #region Ran rate
        public int RANRATE_HEADERS_ROW { get; set; }
        #endregion
    }
}