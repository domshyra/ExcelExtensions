// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Models
{
    public class ImportColumn : ColumnWithSeverity
    {
        public bool IsRequired { get; set; }

        public ImportColumn() : base()
        {
            IsRequired = MissingSeverity == ParseExceptionSeverity.Error;
        }

        public ImportColumn(Column column, bool required) : base(column, required ? ParseExceptionSeverity.Error : ParseExceptionSeverity.Warning)
        {

        }
        public ImportColumn(string modelPropertyName, string displayName, FormatType format, bool required = true, int? decimalPrecision = null) : 
            base(modelPropertyName, displayName, format, required ? ParseExceptionSeverity.Error : ParseExceptionSeverity.Warning, decimalPrecision)
        {

        }
    }
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

    /// <summary>
    /// JUST PARSE
    /// </summary>
    public class InformedImportColumn : ImportColumn
    {
        /// <summary>
        /// used as a pointer for where we are importing and saved if we are looking for a column that doesn't have a letter
        /// <para>This will override the letter and use the number for data</para>
        /// </summary>
        public int ImportColumnNumber { get; set; }

        public InformedImportColumn() : base()
        {

        }


    }


}
