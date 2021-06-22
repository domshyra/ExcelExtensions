// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Models.Columns.Import
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
}
