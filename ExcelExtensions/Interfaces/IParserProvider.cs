using OfficeOpenXml;
using System;

namespace ExcelExtensions.Interfaces
{
    /// <summary>
    /// Provides methods for parsing excel data and returning as in memory data
    /// </summary>
    public interface IParserProvider
    {
        /// <summary>
        /// Provides a <see cref="bool"/> from an <see cref="ExcelRange"/>
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException">If input value or text is null</exception>
        /// <exception cref="FormatException">If trying to convert an invalid text value</exception>
        bool ParseBool(ExcelRange cell);
        /// <summary>
        /// Provides a <see cref="decimal"/> from an <see cref="ExcelRange"/>, removing the $ if it exists.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException">If input value or text is null</exception>
        /// <exception cref="FormatException">If trying to convert an invalid text value</exception>
        decimal? ParseCurrency(ExcelRange cell);
        /// <summary>
        /// Provides a <see cref="DateTime"/> from an <see cref="ExcelRange"/>
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException">If input value or text is null</exception>
        /// <exception cref="FormatException">If trying to convert an invalid text value</exception>
        DateTime? ParseDate(ExcelRange cell);
        /// <summary>
        /// Provides a <see cref="TimeSpan"/> from an <see cref="ExcelRange"/>
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException">If input value or text is null</exception>
        /// <exception cref="FormatException">If trying to convert an invalid text value</exception>
        /// <exception cref="OverflowException">If trying to convert an invalid text value</exception>
        TimeSpan? ParseTimeSpan(ExcelRange cell);
        /// <summary>
        /// Provides a <see cref="decimal"/> from an <see cref="ExcelRange"/>
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException">If input value or text is null</exception>
        /// <exception cref="FormatException">If trying to convert an invalid text value</exception>
        decimal? ParseDecimal(ExcelRange cell);
        /// <summary>
        /// Provides a <see cref="double"/> from an <see cref="ExcelRange"/>
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException">If input value or text is null</exception>
        /// <exception cref="FormatException">If trying to convert an invalid text value</exception>
        double? ParseDouble(ExcelRange cell);
        /// <summary>
        /// Provides an <see cref="int"/> from an <see cref="ExcelRange"/>
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException">If input value or text is null</exception>
        /// <exception cref="FormatException">If trying to convert an invalid text value</exception>
        int? ParseInt(ExcelRange cell);
        /// <summary>
        /// Provides a <see cref="decimal"/> from an <see cref="ExcelRange"/>
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException">If input value or text is null</exception>
        /// <exception cref="FormatException">If trying to convert an invalid text value</exception>
        decimal? ParsePercent(ExcelRange cell);
    }
}