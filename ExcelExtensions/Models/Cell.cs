// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using static Extensions.Enums.Enums;

namespace Extensions.Models
{
    /// <summary>
    /// Represents an excel cell
    /// </summary>
    public class Cell
    {
        /// <summary>
        /// Represents Property name for the mapping object.
        /// </summary>
        public string ModelProperty { get; set; }
        /// <summary>
        /// Represents the cell readable text / display name of the cell
        /// </summary>
        public string DisplayTitle { get; set; }
        /// <summary>
        /// Represents the format type for the excel object
        /// </summary>
        public FormatType Format { get; set; }
        /// <summary>
        /// Represents the location of the label
        /// </summary>
        public string LabelCellLocation { get; set; }
        /// <summary>
        /// Represents the location of the value
        /// </summary>
        public string ValueCellLocation { get; set; }
        /// <summary>
        /// Represents the if the cell is required for import to flag for error or warning if missing
        /// </summary>
        public bool IsRequired { get; set; }
        /// <summary>
        /// Represents the number of decimal places in the output for export
        /// </summary>
        public int? DecimalPrecision { get; set; }

        public Cell()
        {

        }
    }
}
