// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Models
{
    /// <summary>
    /// Represents an excel column
    /// </summary>
    public class ExportColumn : Column
    {
        /// <summary>
        /// Represents the location/letter of the excel column. Ex "C"
        /// <para>Used for exports and import where we know where to look and are not parsing for new col locations</para>
        /// </summary>
        public string? ExportColumnLetter { get; }

        public ExportColumn(int columnNumber) : base()
        {
            ColumnNumber = columnNumber;
        }

        public ExportColumn(IExtensions excelExtensions, string exportColumnLetter) : this(excelExtensions.GetColumnNumber(exportColumnLetter))
        {

        }

        public ExportColumn(string modelPropertyName, string displayName, int columnNumber, FormatType format, int? decimalPrecision = null) : base(modelPropertyName, displayName, format, decimalPrecision)
        {
            ColumnNumber = columnNumber;
        }
    }
}
