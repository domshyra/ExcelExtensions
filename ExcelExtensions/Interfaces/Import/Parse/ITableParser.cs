// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Models;
using OfficeOpenXml;
using System.Collections.Generic;

namespace ExcelExtensions.Interfaces.Import.Parse
{
    /// <summary>
    /// Provides methods to parse a table of data in excel.
    /// </summary>
    /// <typeparam name="T">Object the excel data is being mapped to.</typeparam>
    public interface ITableParser<T>
    {
        ParsedTable<T> ScanForColumnsAndParseTable(List<ImportColumnTemplate> columns, ExcelWorksheet workSheet, int headerRowNumber = 1, int maxScanHeaderRowThreashold = 100);
        ParsedTable<T> ParseTable(List<ImportColumnTemplate> columns, ExcelWorksheet workSheet, int headerRowNumber = 1);
    }
}
