using FilesEditor.Enums;
using System.Collections.Generic;

namespace FilesEditor.Entities
{
    public class SlideToGenerate
    {
        public string OutputFileName;
        public string Title;
        public LayoutTypes LayoutType;
        public List<string> Contents;
    }
}