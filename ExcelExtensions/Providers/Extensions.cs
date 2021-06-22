// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Globals;
using ExcelExtensions.Interfaces;
using ExcelExtensions.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Providers
{
    /// <inheritdoc/>
    public class Extensions : IExtensions
    {
        #region developer exceptions 
        public KeyValuePair<string, ParseException> LogDeveloperException(string worksheetName, ImportColumn column, string cellAddress, string message, string modelPropertyName)
        {
            ParseException parseException = LogDeveloperExceptionParseException(worksheetName, column, cellAddress, message);

            return new KeyValuePair<string, ParseException>(modelPropertyName, parseException);
        }
        public KeyValuePair<int, ParseException> LogDeveloperException(string worksheetName, ImportColumn importColumn, string cellAddress, int rowNumber, string message)
        {
            ParseException parseException = LogDeveloperExceptionParseException(worksheetName, importColumn, cellAddress, message);

            return new KeyValuePair<int, ParseException>(rowNumber, parseException);
        }

        private ParseException LogDeveloperExceptionParseException(string worksheetName, ImportColumn importColumn, string cellAddress, string exeptionMessage)
        {
            //ParseException parseException = new(worksheetName, displayName, cellAddress, null, "An error occurred when trying to set the property info. The error is: " + message)
            ParseException parseException = new(worksheetName, importColumn)
            {
                ColumnLetter = GetColumnLetter(cellAddress),
                Message = $"An error occurred when trying to set the property info. The error is: {exeptionMessage}",
                Row = GetRowNumber(cellAddress),
                ExceptionType = ParseExceptionType.Generic,
                Severity = ParseExceptionSeverity.Error,
            };
            return parseException;
        }

        public KeyValuePair<string, ParseException> LogNullReferenceException(string worksheetName, ImportColumn column, string cellAddress, string modelPropertyName)
        {
            ParseException parseException = LogNullReferenceExceptionParseException(worksheetName, column, cellAddress);

            return new KeyValuePair<string, ParseException>(modelPropertyName, parseException);
        }

        public KeyValuePair<int, ParseException> LogNullReferenceException(string worksheetName, ImportColumn displayName, string cellAddress, int rowNumber)
        {
            ParseException parseException = LogNullReferenceExceptionParseException(worksheetName, displayName, cellAddress);

            return new KeyValuePair<int, ParseException>(rowNumber, parseException);
        }

        private ParseException LogNullReferenceExceptionParseException(string worksheetName, ImportColumn column, string cellAddress)
        {
            ParseException parseException = new(worksheetName, column)
            {
                ColumnLetter = GetColumnLetter(cellAddress),
                Row = GetRowNumber(cellAddress),
                ExceptionType = ParseExceptionType.MissingData,
            };

            return parseException;
        }

        public KeyValuePair<string, ParseException> LogCellException(string worksheetName, ImportColumn column, string cellAddress, string modelPropertyName)
        {
            ParseException parseException = LogCellExceptionParseException(worksheetName, column, cellAddress);

            return new KeyValuePair<string, ParseException>(modelPropertyName, parseException);
        }

        public KeyValuePair<int, ParseException> LogCellException(string worksheetName, ImportColumn column, string cellAddress, int rowNumber)
        {
            ParseException parseException = LogCellExceptionParseException(worksheetName, column, cellAddress);

            return new KeyValuePair<int, ParseException>(rowNumber, parseException);
        }

        private ParseException LogCellExceptionParseException(string worksheetName, ImportColumn column, string cellAddress)
        {
            ParseException parseException = new(worksheetName, column)
            {
                ColumnLetter = GetColumnLetter(cellAddress),
                Row = GetRowNumber(cellAddress),
                ExceptionType = ParseExceptionType.UnexpectedDataType,
            };
            return parseException;
        }
        #endregion
        /// <inheritdoc/>
        public string GetCellAddress(int columnNumber, int rowNumber)
            => $"{GetColumnLetter(columnNumber)}{rowNumber}";

        /// <inheritdoc/>
        public string GetCellAddress(int columnNumber, string rowNumber)
            => $"{GetColumnLetter(columnNumber)}{rowNumber}";

        /// <inheritdoc/>
        public int GetRowNumber(string cellAddress)
        {
            if (string.IsNullOrEmpty(cellAddress.Trim()))
            {
                throw new NullReferenceException($"{nameof(cellAddress)} cannot be {Constants.DefaultIsNullOrEmptyString}.");
            }
            return Convert.ToInt32(Regex.Replace(cellAddress, "[^0-9]+", string.Empty));
        }

        /// <inheritdoc/>
        public int GetColumnNumber(string excelString)
        {
            if (string.IsNullOrEmpty(excelString.Trim()))
            {
                throw new NullReferenceException($"{nameof(excelString)} cannot be {Constants.DefaultIsNullOrEmptyString}.");
            }
            string columnName;
            //Input is an excel Cell not just the column name IE "B12" instead of "B"
            if (excelString.All(char.IsLetter) == false)
            {
                //Remove numbers
                columnName = new string(excelString.Where(char.IsLetter).ToArray());
            }
            //Input is a column name IE "A"
            else
            {
                columnName = excelString;
            }

            int result = 0;

            // Process each letter.
            foreach (char t in columnName)
            {
                result *= 26;
                char letter = t;

                // See if it's out of bounds.
                if (letter < 'A')
                {
                    letter = 'A';
                }
                if (letter > 'Z')
                {
                    letter = 'Z';
                }

                // Add in the value of this letter.
                result += letter - 'A' + 1;
            }
            return result;
        }

        /// <inheritdoc/>
        public (int row, int column) GetCellAddress(string cellAddress)
            => (row: GetRowNumber(cellAddress), column: GetColumnNumber(cellAddress));

        /// <inheritdoc/>
        public string GetColumnLetter(int columnNumber)
        {
            if (columnNumber <= 0)
            {
                return string.Empty;
            }
            int dividend = columnNumber;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
        /// <summary>
        /// Returns the column name from a cell. Ex "C6" returns C.
        /// </summary>
        /// <param name="cell">Cell as a string, ex "C6"</param>
        /// <returns>The column name as an string</returns>
        private static string GetColumnLetter(string cell)
        {
            return Regex.Replace(cell, "[^A-Z]+", string.Empty);
        }
        /// <inheritdoc/>
        public int FindMaxRow(List<Cell> cells)
            => cells.Select(cell => Convert.ToInt32(Regex.Replace(cell.ValueCellLocation, "[^0-9]+", string.Empty))).Concat(new[] { 0 }).Max();


        /// <inheritdoc/>
        public int FindMaxColumn(List<ExportColumn> columns)
            //TODO
            => columns.Select(column => column.ExportColumnNumber).Max();

        /// <inheritdoc/>
        public string AddDecimalPlacesToFormat(Column column, string noDecimals)
            => AddDecimalPlacesToFormat(column.DecimalPrecision, noDecimals);

        /// <inheritdoc/>
        public string AddDecimalPlacesToFormat(Cell cell, string noDecimals)
            => AddDecimalPlacesToFormat(cell.DecimalPrecision, noDecimals);

        /// <inheritdoc/>
        public string AddDecimalPlacesToFormat(int? decimalPrecision, string noDecimals)
        {
            for (int i = 0; i < decimalPrecision; i++)
            {
                if (i == 0)
                {
                    noDecimals += ".";
                }
                noDecimals += "0";
            }

            return noDecimals;
        }

        /// <inheritdoc/>
        public string GetExportModelPropertyNameAndDisplayName(Type objType, string modelPropertyName, out string displayNameAsTitleString)
        {
            try
            {
                //TODO make custome attribute and read info from here
                MemberInfo property = objType.GetProperty(modelPropertyName);

                TextInfo usEnglishTextInfo = new CultureInfo("en-US", false).TextInfo;

                DisplayAttribute attribute = property.GetCustomAttribute<DisplayAttribute>();

                string displayName = property.GetCustomAttribute<DisplayAttribute>()?.Name;
                if (string.IsNullOrEmpty(displayName))
                {
                    displayNameAsTitleString = usEnglishTextInfo.ToTitleCase(modelPropertyName);
                }
                else
                {
                    displayNameAsTitleString = usEnglishTextInfo.ToTitleCase(displayName);
                }


                return modelPropertyName;
            }
            catch (NullReferenceException)
            {
                throw;
            }
        }
    }
}
