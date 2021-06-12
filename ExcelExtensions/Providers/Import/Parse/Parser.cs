// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Globals;
using ExcelExtensions.Interfaces.Import.Parse;
using OfficeOpenXml;
using System;

namespace ExcelExtensions.Providers.Import.Parse
{
    /// <inheritdoc/>
    public class Parser : IParser
    {
        /// <inheritdoc/>
        public decimal? ParsePercent(ExcelRange cell)
        {
            string valueTypeInExcel = Constants.DecimalExceptionTypeString;
            try
            {
                CheckIfCellIsNullOrEmpty(cell, valueTypeInExcel);
            }
            catch (NullReferenceException)
            {
                throw;
            }

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
                throw new FormatException($"{Constants.DefaultCannotString} \"{cell.Text}\" to {valueTypeInExcel}.");
            }
        }

        /// <inheritdoc/>
        public int? ParseInt(ExcelRange cell)
        {
            string valueTypeInExcel = Constants.IntExceptionTypeString;
            try
            {
                CheckIfCellIsNullOrEmpty(cell, valueTypeInExcel);
            }
            catch (NullReferenceException)
            {
                throw;
            }

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
                throw new FormatException($"{Constants.DefaultCannotString} \"{cell.Text}\" to {valueTypeInExcel}.");
            }
        }

        /// <inheritdoc/>
        public decimal? ParseDecimal(ExcelRange cell)
        {
            string valueTypeInExcel = Constants.DecimalExceptionTypeString;
            try
            {
                CheckIfCellIsNullOrEmpty(cell, valueTypeInExcel);
            }
            catch (NullReferenceException)
            {
                throw;
            }

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
                throw new FormatException($"{Constants.DefaultCannotString} \"{cell.Text}\" to {valueTypeInExcel}.");
            }
        }

        /// <inheritdoc/>
        public double? ParseDouble(ExcelRange cell)
        {
            string valueTypeInExcel = Constants.DoubleExceptionTypeString;
            try
            {
                CheckIfCellIsNullOrEmpty(cell, valueTypeInExcel);
            }
            catch (NullReferenceException)
            {
                throw;
            }

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
                throw new FormatException($"{Constants.DefaultCannotString} \"{cell.Text}\" to {valueTypeInExcel}.");
            }
        }

        /// <inheritdoc/>
        public DateTime? ParseDate(ExcelRange cell)
        {
            string valueTypeInExcel = Constants.DateTimeExceptionTypeString;
            try
            {
                CheckIfCellIsNullOrEmpty(cell, valueTypeInExcel);
            }
            catch (NullReferenceException)
            {
                throw;
            }

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
                throw new FormatException($"{Constants.DefaultCannotString} \"{cell.Text}\" to {valueTypeInExcel}.");
            }

        }

        /// <inheritdoc/>
        public TimeSpan? ParseTimeSpan(ExcelRange cell)
        {
            string valueTypeInExcel = Constants.TimeSpanExceptionTypeString;
            try
            {
                CheckIfCellIsNullOrEmpty(cell, valueTypeInExcel);
            }
            catch (NullReferenceException)
            {
                throw;
            }


            //https://docs.microsoft.com/en-us/dotnet/api/system.timespan.tryparseexact?view=net-5.0
            //https://docs.microsoft.com/en-us/dotnet/api/system.timespan.parseexact?view=net-5.0
            //https://docs.microsoft.com/en-us/dotnet/standard/base-types/composite-formatting
            //https://docs.microsoft.com/en-us/dotnet/standard/base-types/custom-timespan-format-strings
            //https://docs.microsoft.com/en-us/dotnet/standard/base-types/standard-timespan-format-strings
            //well lets figure this out later

            try
            {
                TimeSpan interval = TimeSpan.ParseExact(cell.Value?.ToString(), "c", null);

                return interval;
            }
            catch (FormatException)
            {
                Console.WriteLine("{0}: Bad Format", cell.Value);
                //Throw the exception to the calling method
            }
            catch (OverflowException)
            {
                Console.WriteLine("{0}: Out of Range", cell.Value);
            }

            try
            {
                TimeSpan interval = TimeSpan.ParseExact(cell.Text?.ToString(), "c", null);

                return interval;
            }
            catch (FormatException)
            {
                Console.WriteLine("{0}: Bad Format", cell.Text);
                //Throw the exception to the calling method
                throw new FormatException($"{Constants.DefaultCannotString} \"{cell.Text}\" to {valueTypeInExcel}.");
            }
            catch (OverflowException)
            {
                Console.WriteLine("{0}: Out of Range", cell.Text);
                throw new FormatException($"{Constants.DefaultCannotString} \"{cell.Text}\" to {valueTypeInExcel}.");
            }
        }

        /// <inheritdoc/>
        public bool ParseBool(ExcelRange cell)
        {
            bool? returnValue = null;
            string valueTypeInExcel = Constants.BoolExceptionTypeString;
            try
            {
                CheckIfCellIsNullOrEmpty(cell, valueTypeInExcel);
            }
            catch (NullReferenceException)
            {
                throw;
            }

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
                throw new FormatException($"{Constants.DefaultCannotString} \"{cell.Text}\" to {valueTypeInExcel}.");
            }

            return (bool)returnValue;
        }

        /// <inheritdoc/>
        public decimal? ParseCurrency(ExcelRange cell)
        {
            string valueTypeInExcel = Constants.DecimalExceptionTypeString;

            try
            {
                CheckIfCellIsNullOrEmpty(cell, valueTypeInExcel);
            }
            catch (NullReferenceException)
            {
                throw;
            }

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
                throw new FormatException($"{Constants.DefaultCannotString} \"{cell.Text}\" to {valueTypeInExcel}.");
            }
        }

        /// <inheritdoc/>
        //TODO: add unit test for this one
        public string ParseString(ExcelRange cell)
        {
            string valueTypeInExcel = Constants.StringExceptionTypeString;
            try
            {
                CheckIfCellIsNullOrEmpty(cell, valueTypeInExcel);
            }
            catch (NullReferenceException)
            {
                throw;
            }

            try
            {
                string cellVal = cell.Value.ToString().Trim();

                if (string.IsNullOrEmpty(cellVal) && !string.IsNullOrEmpty(cell.Formula))
                {
                    //We have a formula which is setting the field to blank instead of null
                    throw new NullReferenceException();
                }

                return cellVal;
            }

            catch (NullReferenceException)
            {
                throw;
            }
        }

        /// <summary>
        /// Provides a setter for booleans in excel
        /// </summary>
        /// <param name="workingText"></param>
        /// <returns></returns>
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

        /// <inheritdoc/>
        public void CheckIfCellIsNullOrEmpty(ExcelRange cell, string valueTypeInExcel)
        {
            string nullErrorMsg = $"{Constants.DefaultCannotString} {Constants.DefaultIsNullOrEmptyString} to {valueTypeInExcel}.";

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

    }
}
