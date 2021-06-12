// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using Extensions.Models;
using OfficeOpenXml;
using System.Collections.Generic;
using static Extensions.Enums.Enums;

namespace Extensions.Interfaces.Export
{
    /// <summary>
    /// Provides export methods to an <see cref="ExcelWorksheet"/> 
    /// </summary>
    public interface IExporter
    {
        /// <summary>
        /// Export a table to an <see cref="ExcelWorksheet"/>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheet"></param>
        /// <param name="rows"></param>
        /// <param name="columnDefinitions"></param>
        /// <param name="style"></param>
        /// <param name="headerRow"></param>
        void ExportTable<T>(ref ExcelWorksheet sheet, List<T> rows, List<Column> columnDefinitions, Style style, int headerRow = 1);
        /// <summary>
        /// Export list of columns to an <see cref="ExcelWorksheet"/>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheet"></param>
        /// <param name="rows"></param>
        /// <param name="columns"></param>
        /// <param name="headerRow"></param>
        /// <param name="displayNameAdditionalText"></param>
        void ExportColumns<T>(ref ExcelWorksheet sheet, List<T> rows, List<Column> columns, int headerRow = 1, string displayNameAdditionalText = null);
        /// <summary>
        /// Format the column of an <see cref="ExcelWorksheet"/>
        /// </summary>
        /// <param name="column"></param>
        /// <param name="sheet"></param>
        void FormatColumn(Column column, ref ExcelWorksheet sheet);
        /// <summary>
        /// Format the column of an <see cref="ExcelWorksheet"/>
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="column"></param>
        /// <param name="formatter"></param>
        /// <param name="decimalPrecision"></param>
        void FormatColumn(ref ExcelWorksheet sheet, string column, FormatType formatter, int? decimalPrecision = null);
        /// <summary>
        /// Format the column range of an <see cref="ExcelWorksheet"/>
        /// </summary>
        /// <param name="itemcodeSheet"></param>
        /// <param name="startColumn"></param>
        /// <param name="endColumn"></param>
        /// <param name="format"></param>
        /// <param name="decimalPrecision"></param>
        void FormatColumnRange(ExcelWorksheet itemcodeSheet, string startColumn, string endColumn, FormatType format, int? decimalPrecision = null);
        /// <summary>
        /// Style the table header row by the max column in columns of an <see cref="ExcelWorksheet"/>
        /// </summary>
        /// <param name="columns"></param>
        /// <param name="sheet"></param>
        /// <param name="style"></param>
        /// <param name="startrow"></param>
        void StyleTableHeaderRow(List<Column> columns, ref ExcelWorksheet sheet, Style style, int? startrow = null);
    }
}