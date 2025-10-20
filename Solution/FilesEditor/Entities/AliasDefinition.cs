namespace FilesEditor.Entities
{
    internal class AliasDefinition
    {
        public AliasDefinition(string rawValue, string newValue)
        {
            RawValue = rawValue;
            NewValue = newValue;

            var isRegularExpression = rawValue.Contains("*");
            IsRegularExpression = isRegularExpression;
        }
        public string RawValue { get; private set; }
        public string NewValue { get; private set; }
        public bool IsRegularExpression { get; private set; }
    }
}
