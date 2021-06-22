// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Models
{
    public class ColumnWithSeverity : Column
    {
        /// <summary>
        /// Represent severity if this column is not found while parsing 
        /// </summary>
        public ParseExceptionSeverity MissingSeverity { get; set; }

        public ColumnWithSeverity()
        {

        }
        public ColumnWithSeverity(string modelPropertyName, string displayName, FormatType format, ParseExceptionSeverity missingSeverity = ParseExceptionSeverity.Error, int? decimalPrecision = null) :
            base(modelPropertyName, displayName, format, decimalPrecision)
        {
            MissingSeverity = missingSeverity;
        }
        public ColumnWithSeverity(Column column, ParseExceptionSeverity missingSeverity = ParseExceptionSeverity.Error) : this(column.ModelProperty, column.DisplayName, column.Format, missingSeverity, column.DecimalPrecision)
        {
        }

    }

}
