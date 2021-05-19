// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using ExcelExtensions.Models;
using ExcelExtensions.Parsers;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using WebApplication.Models;
using static ExcelExtensions.Enums.Enums;

namespace WebApplication.Providers
{
    public class SampleProvider
    {
        private readonly IExcelExtensionsProvider _excelExtensionsProvider;
        private readonly IFileAndSheetValidatior _validationProvider;

        private readonly TableParser<SampleTableModel> _tableParser;

        public SampleProvider(IExcelExtensionsProvider excelUtilityProvider, IFileAndSheetValidatior validationProvider)
        {
            _excelExtensionsProvider = excelUtilityProvider;
            _validationProvider = validationProvider;

            _tableParser = new TableParser<SampleTableModel>(_excelExtensionsProvider);
        }


        public Dictionary<int, SampleTableModel> SearchAndImportTable(IFormFile file, ImportViewModel model, string password)
        {
            try
            {
                ExcelPackage package = _validationProvider.GetExcelPackage(file, password);

                ExcelWorksheet sheet = _validationProvider.GetExcelWorksheet(package, "Sheet1", out ParseException exception);

                if (sheet == null)
                {
                    //Cannot find sheet throw error to the controller
                    throw new ImportException(exception);
                }

                ParsedTable<SampleTableModel> parsedTable = _tableParser.ScanForColumnsAndParseTable(GetTableColumnTemplates(), sheet);

                if (parsedTable.Errors.Any(x => x.Value.Severity == ParseExceptionSeverity.Error))
                {
                    //If we encountered any errors in the parse lets throw them to the controller
                    throw new ImportException(model.Exceptions);
                }

                return parsedTable.Rows;

            }
            catch (ImportException)
            {
                //Cannot find an excel file
                throw;
            }
        }


        private static List<ImportColumnTemplate> GetTableColumnTemplates()
        {
            return new List<ImportColumnTemplate>()
            {
              
            };
        }
        private List<Column> SampleTableColumns()
        {
            int i = 1;
            Type objType = typeof(SampleTableModel);
            List<Column> excelColumnDefinitionArray = new()
            {
                //                        PropertyName          Display name            col location nam                            format
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.Text), _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.String),
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.Date), _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.Date),
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.DateAsText), _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.Date),
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.DateAsGeneral), _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.Date),
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.Percent), _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.Percent),
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.PercentAsText), _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.Percent),
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.PercentAsNumber), _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.Percent),
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.BoolAsYESNO), _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.Bool),
                new Column("BoolAsTrueFalse", "Bool As True False", _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.Bool),
                new Column("BoolAs10", "Bool As 1 0", _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.Bool),
                new Column("Currency", "Currency", _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.Currency),
                new Column("CurrencyAsText", "Currency As Text", _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.Currency),
                new Column("CurrencyAsGeneral", "Currency As General", _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.Currency),
                new Column("Decimal", "Decimal", _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.Decimal),
                new Column("DecimalAsText", "Decimal AsText", _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.Decimal),
                new Column("RequiredText", "Required Text", _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.String),
                new Column("OptionalText", "Optional Text", _excelExtensionsProvider.GetColumnLetter(i++), ExcelFormatType.String),
            };

            return excelColumnDefinitionArray;
        }
        private List<Column> InitStaticTable()
        {
            int i = 1;
            Type objType = typeof(SampleTableModel);
            List<Column> excelColumnDefinitionArray = new List<Column>
            {
                new Column(_excelExtensionsProvider.GetExportModelPropertyNameAndDisplayName(objType, nameof(SampleTableModel.Text), out string textTitle),
                                          textTitle,
                                          _excelExtensionsProvider.GetColumnLetter(i++),
                                          ExcelFormatType.String),
                new Column(_excelExtensionsProvider.GetExportModelPropertyNameAndDisplayName(objType, nameof(SampleTableModel.Date), out string dateTitle),
                                          dateTitle,
                                          _excelExtensionsProvider.GetColumnLetter(i++),
                                          ExcelFormatType.Date),
                new Column(_excelExtensionsProvider.GetExportModelPropertyNameAndDisplayName(objType, nameof(SampleTableModel.Currency), out string currencyTitle),
                                          currencyTitle,
                                          _excelExtensionsProvider.GetColumnLetter(i++),
                                          ExcelFormatType.Currency),
                new Column(_excelExtensionsProvider.GetExportModelPropertyNameAndDisplayName(objType, nameof(SampleTableModel.RequiredText), out string requiredTextTitle),
                                          requiredTextTitle,
                                          _excelExtensionsProvider.GetColumnLetter(i++),
                                          ExcelFormatType.String),
            };

            return excelColumnDefinitionArray;
        }
    }
}
