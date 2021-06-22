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
        /// <para>If not options we will add the display name</para>
        /// </summary>
        /// <param name="modelPropertyName"></param>
        /// <param name="displayName"></param>
        /// <param name="format"></param>
        /// <param name="required"></param>
        /// <param name="columnOptions"></param>
        /// <param name="decimalPrecision"></param>
        public UninformedImportColumn(string modelPropertyName, string displayName, FormatType format, bool required = true, List<string> columnOptions = null, int? decimalPrecision = null) : base(modelPropertyName, displayName, format, required, decimalPrecision)
        {
            DisplayNameOptions = columnOptions ?? new List<string>() { displayName };
            MissingSeverity = required ? ParseExceptionSeverity.Error : ParseExceptionSeverity.Warning;
        }

        /// <summary>
        /// Adds the header name/title from the column 
        /// </summary>
        /// <param name="column"></param>
        /// <param name="required"></param>
        /// <param name="columnOptions">Optional names to search for</param>
        public UninformedImportColumn(Column column, bool required, List<string> columnOptions = null) : this(column.ModelProperty, column.DisplayName, column.Format, required, columnOptions, column.DecimalPrecision)
        {
        }
    }

}
