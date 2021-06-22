// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Models.Columns.Import;

namespace ExcelExtensions.Models.Columns
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
        public int ColumnNumber { get; set; }

        public InformedImportColumn() : base()
        {

        }

        public InformedImportColumn(UninformedImportColumn column, int columnNumber) : base(column, column.IsRequired)
        {
            ColumnNumber = columnNumber;
        }
    }
}
