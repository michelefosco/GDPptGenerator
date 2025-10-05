using System.Collections.Generic;

namespace FilesEditor.Entities
{
    public class FilterItems
    {
        public string TableName { get; set; }
        public string FieldName { get; set; }
        public List<string> Values { get; set; }
        public List<string> SelectedValues { get; set; }
    }
}
