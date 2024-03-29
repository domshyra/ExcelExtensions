﻿// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

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
        /// <summary>
        /// Represents the location/letter of the excel column. Ex "C"
        /// </summary>
        public string ColumnLetter { get; set; }
        /// <summary>
        /// Represents property name for the mapping object.
        /// </summary>
        public string ModelProperty { get; set; }
        /// <summary>
        /// Represents the human readable column header name/title
        /// </summary>
        public string HeaderTitle { get; set; }
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
        public Column(string modelPropertyName, string headerTitle, string columnLetter, FormatType format, int? decimalPrecision = null)
        {
            ColumnLetter = columnLetter;
            ModelProperty = modelPropertyName;
            HeaderTitle = headerTitle;
            Format = format;
            DecimalPrecision = decimalPrecision;
        }
        public Column(IExtensions _excelExtensions, Type modelType, string modelPropertyName, FormatType format, string columnLetter = null, int? decimalPrecision = null)
        {
            _excelExtensions.GetExportModelPropertyNameAndDisplayName(modelType, modelPropertyName, out string textTitle);
            if (columnLetter != null)
            {
                ColumnLetter = columnLetter;
            }
            ModelProperty = modelPropertyName;
            HeaderTitle = textTitle;
            Format = format;
            DecimalPrecision = decimalPrecision;
        }
        public Column(IExtensions _excelExtensions, Type modelType, string modelPropertyName, FormatType format, int columnLetterAsInt, int? decimalPrecision = null)
        {
            _excelExtensions.GetExportModelPropertyNameAndDisplayName(modelType, modelPropertyName, out string textTitle);
            ColumnLetter = _excelExtensions.GetColumnLetter(columnLetterAsInt);
            ModelProperty = modelPropertyName;
            HeaderTitle = textTitle;
            Format = format;
            DecimalPrecision = decimalPrecision;
        }
    }
}
