using System;

//TODO
namespace ExcelExtensions.Models
{
    /// <summary>
    /// Represents an exception when using the parse functions of EE
    /// </summary>
    public class ImportException : Exception
    {
        public ImportException() : base() { }
        public ImportException(string message) : base(message) { }
        public ImportException(string message, Exception inner) : base(message, inner) { }
    }
}
