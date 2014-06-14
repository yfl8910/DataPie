namespace Kent.Boogaart.KBCsv.Internal
{
    using System.Text;

    internal static class Constants
    {
        public const char DefaultValueSeparator = ',';
        public const char DefaultValueDelimiter = '"';
        public const char Space = ' ';
        public const char Tab = '\t';
        public const char CR = (char)0x0d;
        public const char LF = (char)0x0a;

        public static readonly Encoding DefaultEncoding = Encoding.UTF8;
    }
}
