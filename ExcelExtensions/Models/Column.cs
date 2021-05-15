// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Models
{
    /// <summary>
    /// Represents an excel column
    /// </summary>
    public class Column
    {
        /// <summary>
        /// Represents the location/letter of the excel column. Ex "C"
        /// </summary>
        public string ColumnLetter { get; set; }
        /// <summary>
        /// Represents property name for the mapping object.
        /// </summary>
        public string ModelProperty { get; set; }
        /// <summary>
        /// Represents the human readable column header name
        /// </summary>
        public string HeaderName { get; set; }
        /// <summary>
        /// Represents the format type for the excel object
        /// </summary>
        public ExcelFormatType Format { get; set; }
        /// <summary>
        /// Represents the number of decimal places in the output for export
        /// </summary>
        public int? DecimalPrecision { get; set; }

        public Column()
        {

        }
    }
}
