// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using ExcelExtensions.Interfaces.Export;
using ExcelExtensions.Models.Columns;
using ExcelExtensions.Models.Columns.Export;
using ExcelExtensions.Models.Export;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Providers.Export
{
    public class Exporter : IExporter
    {
        private readonly IExtensions _excelProvider;

        public Exporter(IExtensions excelProvider)
        {
            _excelProvider = excelProvider;
        }

        /// <inheritdoc/>
        public void ExportTable<T>(ref ExcelWorksheet sheet, List<T> rows, List<Column> columnDefinitions, Style style, int headerRow = 1)
        {
            ExportColumns(ref sheet, rows, columnDefinitions, headerRow);

            StyleTableHeaderRow(columnDefinitions, ref sheet, style, headerRow);
        }

        /// <inheritdoc/>
        public void ExportColumns<T>(ref ExcelWorksheet sheet, List<T> rows, List<Column> columns, int headerRow = 1, string displayNameAdditionalText = null)
        {
            Type exportModelType = rows.GetType().GetGenericArguments().Single();

            foreach (ExportColumn column in columns)
            {
                int rowNumber = headerRow;
                rowNumber = SetTableHeaderRowNumber(sheet, displayNameAdditionalText, column, rowNumber);

                foreach (var row in rows)
                {
                    try
                    {
                        PropertyInfo propertyInfo = exportModelType.GetProperty(column.ModelProperty);
                        CheckForNullColumnNumber(column);
                        try
                        {
                            object value = propertyInfo.GetValue(row);
                            sheet.Cells[rowNumber++, (int)column.ColumnNumber].Value = value;
                        }
                        catch (NullReferenceException e)
                        {
                            //propertyInfo info is null
                            throw new NullReferenceException($"column.ModelProperty: [{column.ModelProperty}] failed to get propertyInfo. " +
                                $"Invalid object being cast. Type for {column.ModelProperty} is ${propertyInfo.PropertyType}", e);
                        }
                    }
                    catch (ArgumentNullException e)
                    {
                        //It is a null cell, leave it blank
                        Debug.WriteLine(e.Message);
                    }
                }
                FormatColumn(column, ref sheet);
            }
        }

        /// <summary>
        /// Sets the row number of the table once we find where all the header tiles are located
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="displayNameAdditionalText"></param>
        /// <param name="column"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private static int SetTableHeaderRowNumber(ExcelWorksheet sheet, string displayNameAdditionalText, Column column, int row)
        {
            CheckForNullColumnNumber(column);
            if (string.IsNullOrEmpty(displayNameAdditionalText))
            {
                sheet.Cells[row++, (int)column.ColumnNumber].Value = column.DisplayName;
            }
            else
            {
                sheet.Cells[row++, (int)column.ColumnNumber].Value = $"{column.DisplayName} {displayNameAdditionalText}"; ;
            }

            return row;
        }

        private static void CheckForNullColumnNumber(Column column)
        {
            if (column.ColumnNumber is null)
            {
                throw new NullReferenceException($"Col number cannot be null for property {column.ModelProperty}");
            }
        }

        /// <inheritdoc/>
        public void FormatColumn(Column column, ref ExcelWorksheet sheet)
        {
            CheckForNullColumnNumber(column);
            FormatColumn(ref sheet, (int)column.ColumnNumber, column.Format, column.DecimalPrecision);
        }

        public void FormatColumn(ref ExcelWorksheet sheet, string columnLetter, FormatType formatter, int? decimalPrecision = null)
        {
            FormatColumn(ref sheet, _excelProvider.GetColumnNumber(columnLetter), formatter, decimalPrecision);
        }

        /// <inheritdoc/>
        public void FormatColumn(ref ExcelWorksheet sheet, int columnNumber, FormatType formatter, int? decimalPrecision = null)
        {
            //Styling
            if (formatter == FormatType.String)
            {
                sheet.Column(columnNumber).Width = 21;
            }
            else if (formatter == FormatType.Currency)
            {
                if (decimalPrecision != null)
                {
                    string formatString = "$#,##0";
                    formatString = _excelProvider.AddDecimalPlacesToFormat(decimalPrecision, formatString);
                    sheet.Column(columnNumber).Style.Numberformat.Format = formatString;
                }
                else
                {
                    sheet.Column(columnNumber).Style.Numberformat.Format = "$#,##0.00";
                }
                sheet.Column(columnNumber).Width = 17;
            }
            else if (formatter == FormatType.Date)
            {
                if (DateTimeFormatInfo.CurrentInfo != null)
                {
                    sheet.Column(columnNumber).Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                sheet.Column(columnNumber).Width = 15;
            }
            else if (formatter == FormatType.Duration)
            {

                sheet.Column(columnNumber).Style.Numberformat.Format = "[h]:mm:ss";
                sheet.Column(columnNumber).Width = 15;
            }
            else if (formatter == FormatType.Decimal)
            {
                if (decimalPrecision != null)
                {
                    string formatString = "#,##0";
                    formatString = _excelProvider.AddDecimalPlacesToFormat(decimalPrecision, formatString);
                    sheet.Column(columnNumber).Style.Numberformat.Format = formatString;
                }
                else
                {
                    sheet.Column(columnNumber).Style.Numberformat.Format = "#,##0.0000";
                }
                sheet.Column(columnNumber).Width = 15;
            }
            else if (formatter == FormatType.Int)
            {
                sheet.Column(columnNumber).Style.Numberformat.Format = "0";
                sheet.Column(columnNumber).Width = 13;
            }
            else if (formatter == FormatType.Percent)
            {
                if (decimalPrecision != null)
                {
                    string formatString = "0";
                    formatString = _excelProvider.AddDecimalPlacesToFormat(decimalPrecision, formatString);
                    sheet.Column(columnNumber).Style.Numberformat.Format = formatString += "%";
                }
                else
                {
                    sheet.Column(columnNumber).Style.Numberformat.Format = "0%";
                }
                sheet.Column(columnNumber).Width = 13;
            }
        }

        /// <inheritdoc/>
        public void FormatColumnRange(ExcelWorksheet itemcodeSheet, string startColumn, string endColumn, FormatType format, int? decimalPrecision = null)
        {
            int start = _excelProvider.GetColumnNumber(startColumn);
            int end = _excelProvider.GetColumnNumber(endColumn);
            if (start > end)
            {
                throw new ArgumentOutOfRangeException(nameof(startColumn), $"{nameof(startColumn)} must be less than {nameof(endColumn)}");
            }
            FormatColumnRange(itemcodeSheet, start, end, format, decimalPrecision);
        }
        public void FormatColumnRange(ExcelWorksheet itemcodeSheet, int startColumn, int endColumn, FormatType format, int? decimalPrecision = null)
        {
            
            if (startColumn > endColumn)
            {
                throw new ArgumentOutOfRangeException(nameof(startColumn), $"{nameof(startColumn)} must be less than {nameof(endColumn)}");
            }
            int currenctColumn = startColumn;
            do
            {
                FormatColumn(ref itemcodeSheet, currenctColumn, format, decimalPrecision);

                //if there is only one col, finish after the styling
                if (startColumn == endColumn)
                {
                    break;
                }
                currenctColumn++;
            }
            while (currenctColumn != endColumn);
        }

        /// <inheritdoc/>
        public void StyleTableHeaderRow(List<Column> columns, ref ExcelWorksheet sheet, Style style, int? startrow = null)
        {
            int maxColumn = _excelProvider.FindMaxColumn(columns);
            int row = startrow ?? 1;

            sheet.Cells[row, 1, row, maxColumn].Style.Font.Bold = style.Bold;
            sheet.Cells[row, 1, row, maxColumn].Style.Fill.PatternType = style.PatternType;
            sheet.Cells[row, 1, row, maxColumn].Style.Fill.BackgroundColor.SetColor(style.BackgroundColor);
            sheet.Cells[row, 1, row, maxColumn].Style.Font.Color.SetColor(style.Color);
        }
    }
}
