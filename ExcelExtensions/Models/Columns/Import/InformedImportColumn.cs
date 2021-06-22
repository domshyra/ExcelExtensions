// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Models.Columns.Import
{
    /// <summary>
    /// Works for parse and export
    /// </summary>
    public class InformedImportColumn : ImportColumn
    {
        /// <summary>
        /// used as a pointer for where we are importing and saved if we are looking for a column that doesn't have a letter
        /// <para>This will override the letter and use the number for data</para>
        /// </summary>
        public new int ColumnNumber { get; set; }

        public InformedImportColumn() : base()
        {

        }

        public InformedImportColumn(UninformedImportColumn column, int columnNumber) : base(column, column.IsRequired)
        {
            ColumnNumber = columnNumber;
        }

        public InformedImportColumn(string modelPropertyName, string displayName, int columnNumber, FormatType format, bool required = true, int? decimalPrecision = null) :
            base(modelPropertyName, displayName, format, required, decimalPrecision)
        {
            ColumnNumber = columnNumber;
        }
    }
}
