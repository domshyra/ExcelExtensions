using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Models
{
    /// <summary>
    /// Represents the exception as it occurs in the parse. 
    /// </summary>
    public class ParseException
    {
        /// <summary>
        /// Represents the sheet that the exception occurred
        /// </summary>
        public string Sheet { get; set; }
        public ExcelFormatType? FormatType { get; set; }
        /// <summary>
        /// Represents the column that the exception occurred
        /// </summary>
        public string Column { get; set; }
        /// <summary>
        /// Represents the column header that the exception occurred
        /// </summary>
        public string ColumnHeader { get; set; }
        /// <summary>
        /// Represents the row that the exception occurred
        /// </summary>
        public int Row { get; set; }
        /// <summary>
        /// Represents the message of the exception that occurred
        /// </summary>
        public string Message { get; set; }
        /// <summary>
        /// Represents the <see cref="ParseExceptionSeverity"/> of the exception that occurred
        /// </summary>
        public ParseExceptionSeverity Severity { get; set; }
        /// <summary>
        /// Represents the <see cref="ParseExceptionType"/> of the exception that occurred
        /// </summary>
        public ParseExceptionType Type { get; set; }
    }
}
