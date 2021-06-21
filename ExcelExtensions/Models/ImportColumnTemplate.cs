// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using System.Collections.Generic;
using System.Linq;

namespace ExcelExtensions.Models
{
    //TODO: replace with data attributes
    public class ImportColumn
    {
        /// <summary>
        /// Represent the list of options for the header name
        /// </summary>
        public List<string> ColumnHeaderOptions { get; set; }
        /// <summary>
        /// If required then <see cref="ParseExceptionSeverity.Error"/> if not then warning
        /// </summary>
        public bool IsRequired { get; set; }
        public Column Column { get; set; }
        public ImportColumn()
        {

        }
        /// <summary>
        /// Adds the header name/title from the column 
        /// </summary>
        /// <param name="column"></param>
        /// <param name="required"></param>
        /// <param name="columnOptions">Optional names to search for</param>
        public ImportColumn(Column column, bool required, List<string> columnOptions = null)
        {
            ColumnHeaderOptions = columnOptions ?? new List<string>();
            ColumnHeaderOptions.Add(column.TableHeaderTitle);
            IsRequired = required;
            Column = column;
        }
    }

    /// <summary>
    /// This class has our number populated so that if we have a letter provided we get to have a number
    /// </summary>
    public class ImportColumnWithCellAddress : ImportColumn
    {
        /// <summary>
        /// used as a pointer for where we are importing and saved if we are looking for a column that doesn't have a letter
        /// <para>This will override the letter and use the number for data</para>
        /// </summary>
        public int ImportColumnNumber { get; set; }

        public ImportColumnWithCellAddress() : base()
        {

        }

        public ImportColumnWithCellAddress(IExtensions excelExtensions, ImportColumn template) : base(template.Column, template.IsRequired, template.ColumnHeaderOptions)
        {
            ImportColumnNumber = template.Column.ExportColumnLetter == null ? 0 : excelExtensions.GetColumnNumber(template.Column.ExportColumnLetter);
        }
    }


}
