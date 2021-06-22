// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Models
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false)]
    public class ColumnAttribute : Attribute
    {
        /// <summary>
        /// Represent the list of options for the header name
        /// </summary>
        public List<string> ImportColumnTitleOptions { get; set; }

        /// <summary>
        /// If required then <see cref="ParseExceptionSeverity.Error"/> if not then warning
        /// </summary>
        public bool? IsRequired { get; set; }

        //Might be albe to depricate this
        /// <summary>
        /// Represents the format type for the excel object
        /// </summary>
        public FormatType Format { get; set; }
        /// <summary>
        /// Represents the number of decimal places in the output for export
        /// </summary>
        public int? DecimalPrecision { get; set; }

        public string? ExportColumnLetter { get; set; }

        public int? ExportColumnNumber { get; set; }
        
        public string? ImportColumnLetter { get; set; }

        public int? ImportColumnNumber { get; set; }

        public ColumnAttribute()
        {
            IsRequired = true;
        }

    }
}
