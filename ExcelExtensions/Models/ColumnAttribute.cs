// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Models
{
    /// <summary>
    /// Provides extensions to models for import and export with excel extensions
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, Inherited = false)]
    public class ColumnAttribute : Attribute
    {
        /// <summary>
        /// Represent the list of options for the header name
        /// <para>Default will look at the <see cref="PropertyInfo.Name"/> and the <see cref="DisplayAttribute.Name"/></para>
        /// </summary>
        public List<string> ImportColumnTitleOptions { get; set; }

        /// <summary>
        /// If required then <see cref="ParseExceptionSeverity.Error"/> if not then <see cref="ParseExceptionSeverity.Warning"/>
        /// </summary>
        public bool IsRequired { get; set; }

        //Might be able to deprecate this ????????? who knows yet
        /// <summary>
        /// Represents the format type for the excel object
        /// </summary>
        public FormatType Format { get; set; }
        /// <summary>
        /// Represents the number of decimal places in the output for export
        /// </summary>
        public int? DecimalPrecision { get; set; }

        /// <summary>
        /// Represents the export location
        /// </summary>
        public string? ExportColumnLetter { get; set; }
        /// <summary>
        /// Represents the export location
        /// </summary>
        public int? ExportColumnNumber { get; set; }
        /// <summary>
        /// Represents the import location
        /// </summary>
        public string? ImportColumnLetter { get; set; }
        /// <summary>
        /// Represents the import location
        /// </summary>
        public int? ImportColumnNumber { get; set; }

        public ColumnAttribute()
        {
            IsRequired = true;
        }

        //todo summary
        /// <summary>
        /// 
        /// </summary>
        /// <param name="format"></param>
        /// <param name="titleOptions"></param>
        /// <param name="precision"></param>
        public ColumnAttribute(FormatType format, List<string> titleOptions = null, int? precision = null)
        {
            if (titleOptions is not null)
            {
                ImportColumnTitleOptions = titleOptions;
            }

            if (precision is not null)
            {
                DecimalPrecision = precision;
            }

            Format = format;
            IsRequired = true;
        }
    }
}
