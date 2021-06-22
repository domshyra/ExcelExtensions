// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Models;
using System;
using System.Collections.Generic;

namespace ExcelExtensions.Interfaces
{
    /// <summary>
    /// Provides extensions methods for translating excel addresses, and export helpers.
    /// </summary>
    public interface IExtensions
    {
        /// <summary>
        /// Provides a setter for decimal places for export.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="noDecimals"></param>
        /// <returns></returns>
        string AddDecimalPlacesToFormat(Cell cell, string noDecimals);
        /// <summary>
        /// Provides a setter for decimal places for export.
        /// </summary>
        /// <param name="column"></param>
        /// <param name="noDecimals"></param>
        /// <returns></returns>
        string AddDecimalPlacesToFormat(Column column, string noDecimals);
        /// <summary>
        /// Provides a setter for decimal places for export.
        /// </summary>
        /// <param name="decimalPrecision"></param>
        /// <param name="noDecimals"></param>
        /// <returns></returns>
        string AddDecimalPlacesToFormat(int? decimalPrecision, string noDecimals);
        /// <summary>
        /// Provides the max column used in a sheet based upon columns.
        /// </summary>
        /// <param name="columns">list of ExcelColumnDefinitions</param>
        /// <returns>The max column in the list cells, represented as an <see cref="int"/></returns>
        int FindMaxColumn(List<ExportColumn> columns);
        /// <summary>
        /// Provides the max row used in a cell sheet based upon Cells.
        /// </summary>
        /// <param name="cells">List of ExcelCellDefinitions</param>
        /// <returns>The max row in the list cells, represented as an <see cref="int"/></returns>
        int FindMaxRow(List<Cell> cells);

        /// <summary>
        /// Provides the cell as a <see cref="string"/> from the column number and row number
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <param name="rowNumber"></param>
        /// <returns>Cell as a <see cref="string"/></returns>
        string GetCellAddress(int columnNumber, int rowNumber);
        /// <summary>
        /// Provides the cell as a <see cref="string"/> from the column number and row number
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <param name="rowNumber"></param>
        /// <returns>Cell as a <see cref="string"/></returns>
        string GetCellAddress(int columnNumber, string rowNumber);
        /// <summary>
        /// Provides the row number and column number from an excel cell address
        /// </summary>
        /// <param name="cellAddress"></param>
        /// <returns>Cell as a tuple</returns>
        (int row, int column) GetCellAddress(string cellAddress);
        /// <summary>
        /// Get column Name. Ex 3 returns "C".
        /// </summary>
        /// <param name="columnNumber">Column number as an <see cref="int"/></param>
        /// <returns>the column name as a string from the column number. Returns empty string if null or zero</returns>
        string GetColumnLetter(int columnNumber);
        /// <summary>
        /// Get column number. Ex "J14" returns 10.
        /// </summary>
        /// <param name="excelString">Either an excel cell "G14" or an excel column letter "B"</param>
        /// <returns>Returns the column number of the excel string whether it is an cell format or column letter format</returns>
        /// <exception cref="NullReferenceException">If input value or text is null</exception>
        int GetColumnNumber(string excelString);
        /// <summary>
        /// Sets modelPropertyDisplayName to the Name of the DisplayAttribute formatted as a US English title (Capital letters after every space). 
        /// <para>modelPropertyDisplayName will be set to the ModelPropertyName if DisplayAttribute Name is null.</para>
        /// </summary>
        /// <remarks>This will return the modelPropertyName supplied, so this method can be used inline for a new <see cref="Cell"/> or new <see cref="Column"/></remarks>
        /// <param name="objType">Input/Export model to map an <see cref="Cell"/> or <see cref="Column"/> to</param>
        /// <param name="modelPropertyName">As the return result</param>
        /// <param name="displayNameAsTitleString">The display name formatted as a title string</param>
        /// <returns>The modelPropertyName supplied</returns>
        string GetExportModelPropertyNameAndDisplayName(Type objType, string modelPropertyName, out string displayNameAsTitleString);
        /// <summary>
        /// Get row number from a cell. Ex "C6" returns 6.
        /// </summary>
        /// <param name="cellAddress">Cell as a <see cref="string"/>, ex "C6"</param>
        /// <returns>The row number as an <see cref="int"/></returns>
        /// <exception cref="NullReferenceException">If input value or text is null</exception>
        int GetRowNumber(string cellAddress);







        ///

        KeyValuePair<string, ParseException> LogDeveloperException(string worksheetName, ImportColumn column, string cellAddress, string message, string modelPropertyName);
        KeyValuePair<int, ParseException> LogDeveloperException(string worksheetName, ImportColumn column, string cellAddress, int rowNumber, string message);
        KeyValuePair<string, ParseException> LogNullReferenceException(string worksheetName, ImportColumn column, string cellAddress, string modelPropertyName);
        KeyValuePair<int, ParseException> LogNullReferenceException(string worksheetName, ImportColumn column, string cellAddress, int rowNumber);
        KeyValuePair<int, ParseException> LogCellException(string worksheetName, ImportColumn column, string cellAddress, int rowNumber);
        KeyValuePair<string, ParseException> LogCellException(string worksheetName, ImportColumn column, string cellAddress, string modelPropertyName);

    }
}