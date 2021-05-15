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
        public ImportColumnTemplate(List<string> columnOptions, bool required, Column column)
        {
            ColumnHeaderOptions = columnOptions;
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

        public ImportColumnWithCellAddress(IExcelExtensionsProvider excelExtensions, ImportColumnTemplate template) : base(template.ColumnHeaderOptions, template.IsRequired, template.Column)
        {
            ColumnNumber = excelExtensions.GetColumnNumber(template.Column.ColumnLetter);
        }
    }


}
