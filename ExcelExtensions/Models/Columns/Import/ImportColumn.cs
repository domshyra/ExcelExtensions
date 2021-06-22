// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Models.Columns.Import
{
    public class ImportColumn : Column
    {
        public bool IsRequired { get; set; }
        /// <summary>
        /// Represent severity if this column is not found while parsing 
        /// </summary>
        public ParseExceptionSeverity MissingSeverity { get; set; }

        public ImportColumn() : base()
        {
            IsRequired = true;
            MissingSeverity = ParseExceptionSeverity.Error;
        }
        public ImportColumn(string modelPropertyName, string displayName, FormatType format, bool required = true, int? decimalPrecision = null) :
            base(modelPropertyName, displayName, format, decimalPrecision)
        {
            IsRequired = required;
            MissingSeverity = required ? ParseExceptionSeverity.Error : ParseExceptionSeverity.Warning;
        }
        public ImportColumn(Column column, bool required) : this(column.ModelProperty, column.DisplayName, column.Format, required, column.DecimalPrecision)
        {

        }
    }
}
