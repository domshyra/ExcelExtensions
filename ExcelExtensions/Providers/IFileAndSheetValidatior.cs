using ExcelExtensions.Models;
using OfficeOpenXml;

namespace ExcelExtensions.Interfaces
{
    /// <summary>
    /// 
    /// </summary>
    public interface IFileAndSheetValidatior
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelFile"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        ExcelPackage GetExcelPackage(dynamic excelFile, string password = "");
        /// <summary>
        /// 
        /// </summary>
        /// <param name="package"></param>
        /// <param name="sheetName"></param>
        /// <param name="exception">Returns a <see cref="ParseException"/></param>
        /// <returns></returns>
        ExcelWorksheet GetExcelWorksheet(ExcelPackage package, string sheetName, out ParseException exception);
    }
}