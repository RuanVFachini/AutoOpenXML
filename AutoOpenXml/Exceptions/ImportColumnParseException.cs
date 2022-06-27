using System;

namespace AutoOpenXml.Exceptions
{
    public class ImportColumnParseException : Exception
    {
        public ImportColumnParseException(string message) : base(message)
        {
        }

        public ImportColumnParseException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
