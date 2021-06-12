// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Models;
using OfficeOpenXml;

namespace ExcelExtensions.Interfaces.Import
{
    /// <summary>
    /// Provides methods to get <see cref="ExcelWorksheet"/> and <see cref="ExcelPackage"/>.
    /// </summary>
    public interface IFileAndSheetValidatior
    {
        /// <summary>
        /// Provides an <see cref="ExcelPackage"/> from the excel file.
        /// </summary>
        /// <param name="excelFile"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        /// <exception cref="ImportException">thrown if the file is null or not an excel file.</exception>
        ExcelPackage GetExcelPackage(dynamic excelFile, string password = "");
        /// <summary>
        /// Provides an <see cref="ExcelWorksheet"/> by sheet name. 
        /// </summary>
        /// <param name="package"></param>
        /// <param name="sheetName"></param>
        /// <param name="exception">Returns a <see cref="ParseException"/></param>
        /// <returns>Null if No file found with an <see cref="ParseException"/>, or with a <see cref="ExcelWorksheet"/> and null <paramref name="exception"/></returns>
        ExcelWorksheet GetExcelWorksheet(ExcelPackage package, string sheetName, out ParseException exception);
    }
}