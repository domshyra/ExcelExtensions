// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Models.Columns.Import;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtensions.Models.Columns
{
    /// <summary>
    /// JUST PARSE
    /// </summary>
    public class InformedColumn : ColumnWithSeverity
    {
        /// <summary>
        /// used as a pointer for where we are importing and saved if we are looking for a column that doesn't have a letter
        /// <para>This will override the letter and use the number for data</para>
        /// </summary>
        public int ColumnNumber { get; set; }

        public InformedColumn() : base()
        {

        }

        public InformedColumn(UninformedImportColumn column1, int column2)
        {
        }
    }
}
