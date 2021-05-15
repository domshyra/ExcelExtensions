using System.Collections.Generic;

namespace ExcelExtensions.Models
{
    /// <summary>
    /// Provides parsed results of a table import. 
    /// </summary>
    /// <typeparam name="T">Object of each row parsed in the table.</typeparam>
    public class ParsedTable<T>
    {
        /// <summary>
        /// List of errors and messages found with the key being the excel row number.
        /// </summary>
        public List<KeyValuePair<int, ParseException>> Errors { get; set; }
        /// <summary>
        /// Where key is the excel row number.
        /// </summary>
        public Dictionary<int, T> Rows { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public ParsedTable()
        {
            Errors = new List<KeyValuePair<int, ParseException>>();
            Rows = new Dictionary<int, T>();
        }
    }
}
