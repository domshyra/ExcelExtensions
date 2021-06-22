// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Models;
using ExcelExtensions.Models.Columns;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelExtensions.Providers.Import
{
    public interface IColumnMap
    {
        int GetAttributeColumnNumber(int currentColumnNumber, Dictionary<int, string> columnsUsed, string modelPropertyName, ExcelExtensionsColumnAttribute attribute);
        void GetColumnAttributes(Type modelType, PropertyInfo item, Enums.Enums.FormatType formatType, out string modelPropertyName, out string displayName, out ExcelExtensionsColumnAttribute attribute, out Enums.Enums.FormatType format, out bool required);
        List<ExportColumn> GetExportColumns(Type modelType, Enums.Enums.FormatType formatType = Enums.Enums.FormatType.String, int startColumnNumber = 1);
        List<InAndOutColumn> GetInAndOutColumns(Type modelType, Enums.Enums.FormatType formatType = Enums.Enums.FormatType.String, int startColumnNumber = 1);
        List<InformedImportColumn> GetInformedImportColumns(Type modelType, Enums.Enums.FormatType formatType = Enums.Enums.FormatType.String, int startColumnNumber = 1);
        List<UninformedImportColumn> GetUninformedImportColumns(Type modelType, Enums.Enums.FormatType formatType = Enums.Enums.FormatType.String, int startColumnNumber = 1);
    }
}