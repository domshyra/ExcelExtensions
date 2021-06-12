// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using System.Collections.Generic;
using System.Linq;

namespace ExcelExtensions.Models
{
    public class ImportColumnTemplate
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
        public ImportColumnTemplate()
        {

        }
        /// <summary>
        /// Adds the header name/title from the column 
        /// </summary>
        /// <param name="column"></param>
        /// <param name="required"></param>
        /// <param name="columnOptions">Optional names to search for</param>
        public ImportColumnTemplate(Column column, bool required, List<string> columnOptions = null)
        {
            ColumnHeaderOptions = columnOptions ?? new List<string>();
            ColumnHeaderOptions.Add(column.HeaderTitle);
            IsRequired = required;
            Column = column;
        }
    }

    public class ImportColumnWithCellAddress : ImportColumnTemplate
    {
        public int ColumnNumber { get; set; }
        public string DisplayColumnHeaderNames => ColumnHeaderOptions.Aggregate((a, x) => a + ", " + x);


        public ImportColumnWithCellAddress() : base()
        {

        }

        public ImportColumnWithCellAddress(IExcelExtensionsProvider excelExtensions, ImportColumnTemplate template) : base(template.Column, template.IsRequired, template.ColumnHeaderOptions)
        {
            ColumnNumber = template.Column.ColumnLetter == null ? 0 : excelExtensions.GetColumnNumber(template.Column.ColumnLetter);
        }
    }


}
