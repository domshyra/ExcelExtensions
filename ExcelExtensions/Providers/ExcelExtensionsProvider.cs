﻿// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

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
    public class ExcelExtensionsProvider : IExcelExtensionsProvider
    {
        #region developer exeptpons 
        public KeyValuePair<string, ParseException> LogDeveloperException(string worksheetName, ImportColumnTemplate importColumn, string cellAddress, string message, string modelPropertyName)
        {
            ParseException messageSpecification = LogDeveloperExceptionParseException(worksheetName, importColumn, cellAddress, message);

            return new KeyValuePair<string, ParseException>(modelPropertyName, messageSpecification);
        }
        public KeyValuePair<int, ParseException> LogDeveloperException(string worksheetName, ImportColumnTemplate importColumn, string cellAddress, int rowNumber, string message)
        {
            ParseException messageSpecification = LogDeveloperExceptionParseException(worksheetName, importColumn, cellAddress, message);

            return new KeyValuePair<int, ParseException>(rowNumber, messageSpecification);
        }

        private ParseException LogDeveloperExceptionParseException(string worksheetName, ImportColumnTemplate importColumn, string cellAddress, string exeptionMessage)
        {
            //ParseException messageSpecification = new(worksheetName, displayName, cellAddress, null, "An error occurred when trying to set the property info. The error is: " + message)
            ParseException messageSpecification = new(worksheetName, importColumn.Column)
            {
                ColumnLetter = GetColumnLetter(cellAddress),
                Message = $"An error occurred when trying to set the property info. The error is: {exeptionMessage}",
                Row = GetRowNumber(cellAddress),
                ExceptionType = ParseExceptionType.Generic,
                Severity = ParseExceptionSeverity.Error,
            };
            return messageSpecification;
        }

        public KeyValuePair<string, ParseException> LogNullReferenceException(string worksheetName, ImportColumnTemplate displayName, string cellAddress, string modelPropertyName)
        {
            ParseException messageSpecification = LogNullReferenceExceptionParseException(worksheetName, displayName, cellAddress);

            return new KeyValuePair<string, ParseException>(modelPropertyName, messageSpecification);
        }

        public KeyValuePair<int, ParseException> LogNullReferenceException(string worksheetName, ImportColumnTemplate displayName, string cellAddress, int rowNumber)
        {
            ParseException messageSpecification = LogNullReferenceExceptionParseException(worksheetName, displayName, cellAddress);

            return new KeyValuePair<int, ParseException>(rowNumber, messageSpecification);
        }

        private ParseException LogNullReferenceExceptionParseException(string worksheetName, ImportColumnTemplate column, string cellAddress)
        {
            //ErrorLocation errorLocation = new ErrorLocation(worksheetName, displayName, cellAddress, null, "Missing data");
            ParseException messageSpecification = new(worksheetName, column)
            {
                ColumnLetter = GetColumnLetter(cellAddress),
                Row = GetRowNumber(cellAddress),
                ExceptionType = ParseExceptionType.MissingData,
            };

            return messageSpecification;
        }

        public KeyValuePair<string, ParseException> LogCellException(string worksheetName, ImportColumnTemplate displayName, string cellAddress, string modelPropertyName)
        {
            ParseException messageSpecification = LogCellExceptionParseException(worksheetName, displayName, cellAddress);

            return new KeyValuePair<string, ParseException>(modelPropertyName, messageSpecification);
        }

        public KeyValuePair<int, ParseException> LogCellException(string worksheetName, ImportColumnTemplate displayName, string cellAddress, int rowNumber)
        {
            ParseException messageSpecification = LogCellExceptionParseException(worksheetName, displayName, cellAddress);

            return new KeyValuePair<int, ParseException>(rowNumber, messageSpecification);
        }

        private ParseException LogCellExceptionParseException(string worksheetName, ImportColumnTemplate column, string cellAddress)
        {
            ParseException messageSpecification = new(worksheetName, column)
            {
                ColumnLetter = GetColumnLetter(cellAddress),
                Row = GetRowNumber(cellAddress),
                ExceptionType = ParseExceptionType.DataTypeExpectedError,
            };
            return messageSpecification;
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
        public int FindMaxColumn(List<Column> columns)
            => columns.Select(column => GetColumnNumber(column.ColumnLetter)).Max();

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
                MemberInfo property = objType.GetProperty(modelPropertyName);

                TextInfo usEnglishTextInfo = new CultureInfo("en-US", false).TextInfo;

                DisplayAttribute attribute = property.GetCustomAttribute<DisplayAttribute>();

                string? displayName = property.GetCustomAttribute<DisplayAttribute>()?.Name;
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
