// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using ExcelExtensions.Models.Columns.Import;
using System;
using System.Collections.Generic;
using System.Linq;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Models
{
    
    /// <summary>
    /// SCAN AND PARSE
    /// </summary>
    public class UninformedImportColumn : ImportColumn
    {
        /// <summary>
        /// Represent the list of options for the header name
        /// <para>Used for parseing</para>
        /// </summary>
        public List<string> DisplayNameOptions { get; set; }
        public UninformedImportColumn()
        {

        }


        /// <summary>
        /// Adds the header name/title from the column 
        /// </summary>
        /// <param name="column"></param>
        /// <param name="required"></param>
        /// <param name="columnOptions">Optional names to search for</param>
        public UninformedImportColumn(Column column, bool required, List<string> columnOptions = null) : base(column, required)
        {
            DisplayNameOptions = columnOptions ?? new List<string>();
            DisplayNameOptions.Add(column.DisplayName);
            MissingSeverity = required ? ParseExceptionSeverity.Error : ParseExceptionSeverity.Warning;
        }
    }

}
