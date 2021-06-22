// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Models;
using ExcelExtensions.Models.Columns;
using ExcelExtensions.Models.Columns.Export;
using ExcelExtensions.Models.Columns.Import;
using System;
using System.Collections.Generic;
using System.Reflection;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Interfaces.Import
{
    //todo fix these comments 
    /// <summary>
    /// Provides methods to give to the parse methods from models 
    /// </summary>
    public interface IColumnMap
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="import"></param>
        /// <param name="currentColumnNumber"></param>
        /// <param name="columnsUsed"></param>
        /// <param name="modelPropertyName"></param>
        /// <param name="attribute"></param>
        /// <returns></returns>
        int GetAttributeColumnNumber(bool import, int currentColumnNumber, Dictionary<int, string> columnsUsed, string modelPropertyName, ColumnAttribute attribute);
        /// <summary>
        /// 
        /// </summary>
        /// <param name="modelType"></param>
        /// <param name="item"></param>
        /// <param name="formatType"></param>
        /// <param name="modelPropertyName"></param>
        /// <param name="displayName"></param>
        /// <param name="attribute"></param>
        /// <param name="format"></param>
        /// <param name="required"></param>
        void GetColumnAttributes(Type modelType, PropertyInfo item, FormatType formatType, out string modelPropertyName, out string displayName, out ColumnAttribute attribute, out FormatType format, out bool required);
        /// <summary>
        /// 
        /// </summary>
        /// <param name="modelType"></param>
        /// <param name="formatType"></param>
        /// <param name="startColumnNumber"></param>
        /// <returns></returns>
        List<ExportColumn> GetExportColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1);
        /// <summary>
        /// 
        /// </summary>
        /// <param name="modelType"></param>
        /// <param name="formatType"></param>
        /// <param name="startColumnNumber"></param>
        /// <returns></returns>
        List<InAndOutColumn> GetInAndOutColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1);
        /// <summary>
        /// 
        /// </summary>
        /// <param name="modelType"></param>
        /// <param name="formatType"></param>
        /// <param name="startColumnNumber"></param>
        /// <returns></returns>
        List<InformedImportColumn> GetInformedImportColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1);
        /// <summary>
        /// 
        /// </summary>
        /// <param name="modelType"></param>
        /// <param name="formatType"></param>
        /// <param name="startColumnNumber"></param>
        /// <returns></returns>
        List<UninformedImportColumn> GetUninformedImportColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1);
    }
}