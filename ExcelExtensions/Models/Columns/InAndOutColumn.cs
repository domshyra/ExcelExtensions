// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Models.Columns.Import;
using System.Collections.Generic;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Models.Columns
{
    /// <summary>
    /// Pre defined export location based on colom order, finds colums based on names provided
    /// </summary>
    public class InAndOutColumn : UninformedImportColumn
    {

        /// <summary>
        /// Becasue this will use an uninformed when parsing we will search
        /// </summary>
        public new int ColumnNumber { get; set; }

        public InAndOutColumn() : base()
        {

        }

        public InAndOutColumn(string modelPropertyName, string displayName, int columnNumber, FormatType format, bool required = true, List<string> columnOptions = null, int? decimalPrecision = null) :
            base(modelPropertyName, displayName, format, required, columnOptions, decimalPrecision)
        {
            ColumnNumber = columnNumber;
        }
    }
}
