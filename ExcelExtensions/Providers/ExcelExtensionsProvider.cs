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

namespace ExcelExtensions.Providers
{
    /// <inheritdoc/>
    public class ExcelExtensionsProvider : IExcelExtensionsProvider
    {

        /// <inheritdoc/>
        public string GetCellAddress(int columnNumber, int rowNumber)
            => $"{GetColumnName(columnNumber)}{rowNumber}";

        /// <inheritdoc/>
        public string GetCellAddress(int columnNumber, string rowNumber)
            => $"{GetColumnName(columnNumber)}{rowNumber}";

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
        public string GetColumnName(int columnNumber)
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
