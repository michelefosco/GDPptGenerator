using FilesEditor.Enums;
using System.Collections.Generic;

namespace FilesEditor.Entities
{
    public class InputDataFilters_Items
    {
        public InputDataFilters_Tables Table { get; set; }
        public string FieldName { get; set; }
        public List<string> Values { get; set; }
        public List<string> SelectedValues { get; set; }
    }
}
