// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using System;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Models
{
    /// <summary>
    /// Represents an excel column
    /// </summary>
    public class ExportColumn : Column
    {
        /// <summary>
        /// Represents the location/letter of the excel column. Ex "C"
        /// <para>Used for exports and import where we know where to look and are not parsing for new col locations</para>
        /// </summary>
        [Obsolete]
        public string ExportColumnLetter { get; set; }
        public int ExportColumnNumber { get; set; }
        /// <summary>
        /// Represents property name for the mapping object.
        /// </summary>
        public string ModelProperty { get; set; }
        /// <summary>
        /// Represents the human readable column header name/title
        /// </summary>
        public string TableHeaderTitle { get; set; }
        /// <summary>
        /// Represents the format type for the excel object
        /// </summary>
        public FormatType Format { get; set; }
        /// <summary>
        /// Represents the number of decimal places in the output for export
        /// </summary>
        public int? DecimalPrecision { get; set; }

        public ExportColumn()
        {

        }
        public ExportColumn(string modelPropertyName, string headerTitle, int columnLetterAsInt, FormatType format, int? decimalPrecision = null)
        {
            ExportColumnNumber = columnLetterAsInt;
            ModelProperty = modelPropertyName;
            TableHeaderTitle = headerTitle;
            Format = format;
            DecimalPrecision = decimalPrecision;
        }
        public Column(IExtensions _excelExtensions, Type modelType, string modelPropertyName, FormatType format, string columnLetter = null, int? decimalPrecision = null)
        {
            _excelExtensions.GetExportModelPropertyNameAndDisplayName(modelType, modelPropertyName, out string textTitle);
            if (columnLetter != null)
            {
                ExportColumnNumber = _excelExtensions.GetColumnNumber(columnLetter);
            }
            ModelProperty = modelPropertyName;
            TableHeaderTitle = textTitle;
            Format = format;
            DecimalPrecision = decimalPrecision;
        }
        public Column(IExtensions _excelExtensions, Type modelType, string modelPropertyName, FormatType format, int columnLetterAsInt, int? decimalPrecision = null)
        {
            _excelExtensions.GetExportModelPropertyNameAndDisplayName(modelType, modelPropertyName, out string textTitle);
            ExportColumnNumber = columnLetterAsInt;
            ModelProperty = modelPropertyName;
            TableHeaderTitle = textTitle;
            Format = format;
            DecimalPrecision = decimalPrecision;
        }
    }
}
