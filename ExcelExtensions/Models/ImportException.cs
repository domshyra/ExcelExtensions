using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//TODO
namespace ExcelExtensions.Models
{
    /// <summary>
    /// Represents an exception when using the parse funtions of EE
    /// </summary>
    public class ImportException : Exception
    {
        public ImportException() : base() { }
        public ImportException(string message) : base(message) { }
        public ImportException(string message, Exception inner) : base(message, inner) { }
    }
}
