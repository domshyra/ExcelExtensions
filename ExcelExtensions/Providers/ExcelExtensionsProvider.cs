using ExcelExtensions.Interfaces;
using ExcelExtensions.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ExcelExtensions.Providers
{
    public class ExcelExtensionsProvider : IExcelExtensionsProvider
    {
        private const string DefaultCannotString = "Cannot convert";
        private const string DefaultIsNullOrEmptyString = "null or empty value";

        #region Parse methods
        public decimal? ParsePercent(ExcelRange cell)
        {
            string valueTypeInExcel = "a decimal value";
            CheckIfCellIsNullOrEmpty(cell, valueTypeInExcel);

            try
            {
                bool result = decimal.TryParse(cell.Value?.ToString().Replace("%", string.Empty), out decimal returnValue);

                if (result)
                {
                    return returnValue;
                }
                else
                {
                    return Convert.ToDecimal(cell.Text.Replace("%", string.Empty));
                }
            }
            catch (FormatException)
            {
                //Throw the exception to the calling method
                throw new FormatException($"{DefaultCannotString} \"{cell.Text}\" to {valueTypeInExcel}.");
            }
        }

        public int? ParseInt(ExcelRange cell)
        {
            string valueTypeInExcel = "an int value";
            CheckIfCellIsNullOrEmpty(cell, valueTypeInExcel);

            try
            {
                //try to parse a a decimal so we can convert to int later. No need to fail on "1.0" that is "1"
                bool result = decimal.TryParse(cell.Value?.ToString(), out decimal returnValue);

                if (result)
                {
                    return (int)returnValue;
                }
                else
                {
                    return Convert.ToInt32(cell.Text);
                }
            }
            catch (FormatException)
            {
                //Throw the exception to the calling method
                throw new FormatException($"{DefaultCannotString} \"{cell.Text}\" to {valueTypeInExcel}.");
            }
        }

        public decimal? ParseDecimal(ExcelRange cell)
        {
            string valueTypeInExcel = "a decimal value";
            CheckIfCellIsNullOrEmpty(cell, valueTypeInExcel);

            try
            {
                bool result = decimal.TryParse(cell.Value?.ToString(), out decimal returnValue);

                if (result)
                {
                    return returnValue;
                }
                else
                {
                    return Convert.ToDecimal(cell.Text);
                }
            }
            catch (FormatException)
            {
                //Throw the exception to the calling method
                throw new FormatException($"{DefaultCannotString} \"{cell.Text}\" to {valueTypeInExcel}.");
            }
        }

        public double? ParseDouble(ExcelRange cell)
        {
            string valueTypeInExcel = "a double value";
            CheckIfCellIsNullOrEmpty(cell, valueTypeInExcel);

            try
            {
                bool result = double.TryParse(cell.Value?.ToString(), out double returnValue);

                if (result)
                {
                    return returnValue;
                }
                else
                {
                    return Convert.ToDouble(cell.Text);
                }
            }
            catch (FormatException)
            {
                //Throw the exception to the calling method
                throw new FormatException($"{DefaultCannotString} \"{cell.Text}\" to {valueTypeInExcel}.");
            }
        }

        public DateTime? ParseDate(ExcelRange cell)
        {
            string valueTypeInExcel = "a DateTime value";
            CheckIfCellIsNullOrEmpty(cell, valueTypeInExcel);

            try
            {
                DateTime returnValue = DateTime.Now;

                bool result = DateTime.TryParse(cell.Value?.ToString(), out returnValue);

                if (result)
                {
                    //We have a datetime as a string
                    return returnValue;
                }
                else
                {
                    //The cell value did not pan out as a string, so lets try the text
                    result = DateTime.TryParse(cell.Text, out returnValue);
                }

                if (result)
                {
                    //The cell text worked as a string
                    return returnValue;
                }
                else
                {
                    try
                    {
                        //The cell text did not work, let's try the value as a double as a double.
                        return returnValue = DateTime.FromOADate((double)cell.Value);
                    }
                    catch (Exception)
                    {
                        //The cell vale did not work, let's try the text as a double as a double.
                        return returnValue = DateTime.FromOADate(Convert.ToDouble(cell.Text));
                    }
                }
            }
            catch (FormatException)
            {
                //Throw the exception to the calling method
                throw new FormatException($"{DefaultCannotString} \"{cell.Text}\" to {valueTypeInExcel}.");
            }

        }

        public bool ParseBool(ExcelRange cell)
        {
            bool? returnValue = null;
            string valueTypeInExcel = "a bool value";
            CheckIfCellIsNullOrEmpty(cell, valueTypeInExcel);

            if (!string.IsNullOrEmpty(cell.Value?.ToString()))
            {
                returnValue = SetBool(cell.Value.ToString());
            }

            if (!string.IsNullOrEmpty(cell.Text?.ToString()))
            {
                returnValue = SetBool(cell.Text);
            }

            if (returnValue == null)
            {
                throw new FormatException($"{DefaultCannotString} \"{cell.Text}\" to {valueTypeInExcel}.");
            }

            return (bool)returnValue;
        }

        public decimal? ParseCurrency(ExcelRange cell)
        {
            string valueTypeInExcel = "a decimal value";
            CheckIfCellIsNullOrEmpty(cell, valueTypeInExcel);

            try
            {
                bool result = decimal.TryParse(cell.Value?.ToString().Replace("$", string.Empty), out decimal returnValue);

                if (result)
                {
                    return returnValue;
                }
                else
                {
                    return Convert.ToDecimal(cell.Text.Replace("$", string.Empty));
                }
            }
            catch (FormatException)
            {
                //Throw the exception to the calling method
                throw new FormatException($"{DefaultCannotString} \"{cell.Text}\" to {valueTypeInExcel}.");
            }
        }

        private static bool? SetBool(string workingText)
        {
            string trimmedText = workingText.Trim();
            if (string.Equals("yes", trimmedText, StringComparison.OrdinalIgnoreCase) ||
                string.Equals("true", trimmedText, StringComparison.OrdinalIgnoreCase) ||
                string.Equals("1", trimmedText, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            else if (string.Equals("no", trimmedText, StringComparison.OrdinalIgnoreCase) ||
                     string.Equals("false", trimmedText, StringComparison.OrdinalIgnoreCase) ||
                     string.Equals("0", trimmedText, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return null;
        }

        private static void CheckIfCellIsNullOrEmpty(ExcelRange cell, string valueTypeInExcel)
        {
            string nullErrorMsg = $"{DefaultCannotString} {DefaultIsNullOrEmptyString} to {valueTypeInExcel}.";

            //Lets check to see if we have any values to work with
            if (string.IsNullOrEmpty(cell.Value?.ToString()) && string.IsNullOrEmpty(cell.Text?.ToString()))
            {
                //Nothing to parse
                throw new NullReferenceException(nullErrorMsg);
            }
            else if (string.IsNullOrEmpty(cell.Value?.ToString()) && !string.IsNullOrEmpty(cell.Formula))
            {
                //We have nothing to parse because there is a formula causing a blank cell
                throw new NullReferenceException(nullErrorMsg);
            }
        }

        #endregion

        #region Excel methods

        public string GetCellAddress(int colNumber, int rowNumber)
            => $"{GetColumnName(colNumber)}{rowNumber}";

        public string GetCellAddress(int colNumber, string row)
            => $"{GetColumnName(colNumber)}{row}";

        public int GetRowNumber(string cellAddress)
        {
            if (string.IsNullOrEmpty(cellAddress.Trim()))
            {
                throw new NullReferenceException($"{nameof(cellAddress)} cannot be {DefaultIsNullOrEmptyString}.");
            }
            return Convert.ToInt32(Regex.Replace(cellAddress, "[^0-9]+", string.Empty));
        }

        public int GetColumnNumber(string excelString)
        {
            if (string.IsNullOrEmpty(excelString.Trim()))
            {
                throw new NullReferenceException($"{nameof(excelString)} cannot be {DefaultIsNullOrEmptyString}.");
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

        public (int row, int column) GetCellAddress(string cellAddress) 
            => (row: GetRowNumber(cellAddress), column: GetColumnNumber(cellAddress));

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

        public int FindMaxRow(List<Cell> cells)
            => cells.Select(cell => Convert.ToInt32(Regex.Replace(cell.ValueCellLocation, "[^0-9]+", string.Empty))).Concat(new[] { 0 }).Max();


        public int FindMaxColumn(List<Column> columns)
            => columns.Select(column => GetColumnNumber(column.ColumnLetter)).Max();

        public string AddDecimalPlacesToFormat(Column column, string noDecimals) 
            => AddDecimalPlacesToFormat(column.DecimalPrecision, noDecimals);

        public string AddDecimalPlacesToFormat(Cell cell, string noDecimals) 
            => AddDecimalPlacesToFormat(cell.DecimalPrecision, noDecimals);

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

        #endregion Helpers Methods
    }
}
