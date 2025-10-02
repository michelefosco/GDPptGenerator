using System.Collections.Generic;

namespace FilesEditor.Entities
{
    public class FilterItems
    {
        public string Tabella { get; set; }
        public string Campo { get; set; }
        public List<string> ValoriPossibili { get; set; }
        public List<string> ValoriSelezionati { get; set; }
    }
}
