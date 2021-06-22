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
    public interface IColumnMap
    {
        int GetAttributeColumnNumber(int currentColumnNumber, Dictionary<int, string> columnsUsed, string modelPropertyName, ExcelExtensionsColumnAttribute attribute);
        void GetColumnAttributes(Type modelType, PropertyInfo item, FormatType formatType, out string modelPropertyName, out string displayName, out ExcelExtensionsColumnAttribute attribute, out FormatType format, out bool required);
        List<ExportColumn> GetExportColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1);
        List<InAndOutColumn> GetInAndOutColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1);
        List<InformedImportColumn> GetInformedImportColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1);
        List<UninformedImportColumn> GetUninformedImportColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1);
    }
}