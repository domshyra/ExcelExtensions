// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using System;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Models
{
    /// <summary>
    /// Represents an excel column
    /// </summary>
    public class Column
    {
        private IExtensions excelExtensions;
        private Type objType;
        private string v1;
        private FormatType @string;
        private int v2;

        public string ModelProperty { get; set; }
        /// <summary>
        /// Represents the human readable column header name/title
        /// </summary>
        public string DisplayName { get; set; }
        /// <summary>
        /// Represents the format type for the excel object
        /// </summary>
        public FormatType Format { get; set; }
        /// <summary>
        /// Represents the number of decimal places in the output for export
        /// </summary>
        public int? DecimalPrecision { get; set; }

        public Column()
        {

        }
        public Column(string modelPropertyName, string displayName, FormatType format, int? decimalPrecision = null)
        {
            ModelProperty = modelPropertyName;
            DisplayName = displayName;
            Format = format;
            DecimalPrecision = decimalPrecision;
        }

        public Column(IExtensions excelExtensions, Type objType, string v1, FormatType @string, int v2)
        {
            this.excelExtensions = excelExtensions;
            this.objType = objType;
            this.v1 = v1;
            this.@string = @string;
            this.v2 = v2;
        }
    }
}
