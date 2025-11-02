namespace FilesEditor.Entities
{
    internal class RigaBudgetForecast
    {
        public RigaBudgetForecast( string business, string categoria, double[] columns)
        {
            Business = business; 
            Categoria = categoria;            
            Columns = columns;
        }

        public string Business { get; private set; }
        public string Categoria { get; private set; }        
        public double[] Columns { get; private set; }
    }
}
