using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Models
{
    /// <summary>
    /// Represents the error as it occurs in the parse. 
    /// </summary>
    public class ParseError
    {
        /// <summary>
        /// Represents the sheet the error occurred on
        /// </summary>
        public string Sheet { get; set; }
        public ExcelFormatType? FormatType { get; set; }
        /// <summary>
        /// Represents the column the error occurred on
        /// </summary>
        public string Column { get; set; }
        /// <summary>
        /// Represents the column header we are trying to parse in
        /// </summary>
        public string ColumnHeader { get; set; }
        /// <summary>
        /// Represents the row the error occurred on
        /// </summary>
        public int Row { get; set; }
        /// <summary>
        /// Represents the message of the error that occurred
        /// </summary>
        public string Message { get; set; }
    }
}
