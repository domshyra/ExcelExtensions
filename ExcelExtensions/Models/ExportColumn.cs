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

        public int ExportColumnNumber { get; }

        /// <summary>
        /// Represents the location/letter of the excel column. Ex "C"
        /// <para>Used for exports and import where we know where to look and are not parsing for new col locations</para>
        /// </summary>
        public string? ExportColumnLetter { get; }

        public ExportColumn(int columnNumber) : base()
        {
            ExportColumnNumber = columnNumber;
        }

        public ExportColumn(IExtensions excelExtensions, string exportColumnLetter) : this (excelExtensions.GetColumnNumber(exportColumnLetter))
        {

        }
    }
}
