using ExcelExtensions.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace ExcelExtensions.Interfaces
{
    //TODO add comments to these methods
    /// <summary>
    /// 
    /// </summary>
    public interface IExcelExtensionsProvider
    {
        #region Excel extensions
        string AddDecimalPlacesToFormat(Cell cell, string noDecimals);
        string AddDecimalPlacesToFormat(Column column, string noDecimals);
        string AddDecimalPlacesToFormat(int? decimalPrecision, string noDecimals);
        int FindMaxColumn(List<Column> columns);
        int FindMaxRow(List<Cell> cells);
        string GetCellAddress(int colNumber, int rowNumber);
        string GetCellAddress(int colNumber, string row);
        (int row, int column) GetCellAddress(string cellAddress);
        string GetColumnName(int columnNumber);
        int GetColumnNumber(string excelString);
        string GetExportModelPropertyNameAndDisplayName(Type objType, string modelPropertyName, out string displayNameAsTitleString);
        int GetRowNumber(string cellAddress);
        #endregion

        #region Parsing
        bool ParseBool(ExcelRange cell);
        decimal? ParseCurrency(ExcelRange cell);
        DateTime? ParseDate(ExcelRange cell);
        decimal? ParseDecimal(ExcelRange cell);
        double? ParseDouble(ExcelRange cell);
        int? ParseInt(ExcelRange cell);
        decimal? ParsePercent(ExcelRange cell);
        #endregion
    }
}