// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using Extensions.Globals;
using static Extensions.Enums.Enums;

namespace Extensions.Models
{
    /// <summary>
    /// Represents the exception as it occurs in the parse. 
    /// </summary>
    public class ParseException
    {
        /// <summary>
        /// Represents the sheet that the exception occurred
        /// </summary>
        public string Sheet { get; set; }
        public FormatType? FormatType { get; set; }
        /// <summary>
        /// Represents the column that the exception occurred
        /// </summary>
        public string ColumnLetter { get; set; }
        /// <summary>
        /// Represents the column header that the exception occurred
        /// </summary>
        public string ColumnHeader { get; set; }
        /// <summary>
        /// Represents the row that the exception occurred
        /// </summary>
        public int Row { get; set; }
        /// <summary>
        /// Represents the message of the exception that occurred
        /// </summary>
        public string Message { get; set; }
        /// <summary>
        /// Represents the <see cref="ParseExceptionSeverity"/> of the exception that occurred
        /// </summary>
        public ParseExceptionSeverity Severity { get; set; }
        /// <summary>
        /// Represents the <see cref="ParseExceptionType"/> of the exception that occurred
        /// </summary>
        public ParseExceptionType ExceptionType { get; set; }
        /// <summary>
        /// String print out of the expected type
        /// </summary>
        public string ExpectedDateType => FormatType switch
        {
            (Enums.Enums.FormatType.Bool) => Constants.BoolTypeExpectedString,
            (Enums.Enums.FormatType.Currency) => Constants.CurrencyTypeExpectedString,
            (Enums.Enums.FormatType.Date) => Constants.DateTimeTypeExpectedString,
            (Enums.Enums.FormatType.Decimal) => Constants.DecimalTypeExpectedString,
            (Enums.Enums.FormatType.DecimalList) => Constants.DecimalListTypeExpectedString,
            (Enums.Enums.FormatType.Double) => Constants.DoubleTypeExpectedString,
            (Enums.Enums.FormatType.Duration) => Constants.TimeSpanTypeExpectedString,
            (Enums.Enums.FormatType.Int) => Constants.IntTypeExpectedString,
            (Enums.Enums.FormatType.Percent) => Constants.PercentTypeExpectedString,
            (Enums.Enums.FormatType.String) => Constants.StringTypeExpectedString,
            (Enums.Enums.FormatType.StringList) => Constants.StringListTypeExpectedString,
            _ => "",
        };

        public ParseException()
        {

        }
        public ParseException(string sheetName)
        {
            Sheet = sheetName;
        }
        public ParseException(string sheetName, Column column) : this (sheetName)
        {
            ColumnHeader = column.ModelProperty;
            //ColumnHeader = column.ColumnHeader; //todo
            FormatType = column.Format;
        }
        public ParseException(string sheetName, ImportColumnTemplate column) : this (sheetName, column.Column)
        {
            Severity = (column.IsRequired) ? ParseExceptionSeverity.Error : ParseExceptionSeverity.Warning;
        }
    }
}
