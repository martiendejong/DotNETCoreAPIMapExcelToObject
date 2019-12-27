namespace ExcelToObject
{
    public class MappedField
    {
        public readonly string SourceName;

        public readonly string DestinationName;

        public MappedField(string source, string destination)
        {
            SourceName = source;
            DestinationName = destination;
        }
    }
}
