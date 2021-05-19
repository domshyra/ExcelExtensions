// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using ExcelExtensions.Models;
using ExcelExtensions.Parsers;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using WebApplication.Interfaces;
using WebApplication.Models;
using static ExcelExtensions.Enums.Enums;

namespace WebApplication.Providers
{
    public class SampleProvider : ISampleProvider
    {
        private readonly IExcelExtensionsProvider _excelExtensionsProvider;
        private readonly IFileAndSheetValidatior _validationProvider;

        private readonly TableParser<SampleTableModel> _tableParser;

        public SampleProvider(IExcelExtensionsProvider excelUtilityProvider, IFileAndSheetValidatior validationProvider, IParserProvider parserProvider)
        {
            _excelExtensionsProvider = excelUtilityProvider;
            _validationProvider = validationProvider;

            _tableParser = new TableParser<SampleTableModel>(_excelExtensionsProvider, parserProvider);
        }


        public Dictionary<int, SampleTableModel> ScanForColumnsAndParseTable(IFormFile file, ImportViewModel model, string password)
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


        private List<ImportColumnTemplate> GetTableColumnTemplates()
        {
            var cols = SampleTableColumns();

            return new List<ImportColumnTemplate>()
            {
              //                       Col                                                             Req?   
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.Text)),                 true),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.Date)),                 true),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.DateAsText)),           true),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.DateAsGeneral)),        false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.Percent)),              true),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.PercentAsText)),        false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.PercentAsNumber)),      false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.BoolAsYESNO)),          false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.BoolAsTrueFalse)),      false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.BoolAs10)),             false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.Currency)),             false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.CurrencyAsText)),       false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.CurrencyAsGeneral)),    false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.Decimal)),              false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.DecimalAsText)),        false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.RequiredText)),         true),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.OptionalText)),         false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.ListOfStrings)),        false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.ListOfStrings)),        false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.ListOfStrings)),        false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.ListOfDecimals)),       false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.ListOfDecimals)),       false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.ListOfDecimals)),       false),
            };
        }

        private static Column GetColumn(List<Column> cols, string modelPropertyName)
        {
            return cols.First(x => x.ModelProperty == modelPropertyName);
        }
        private List<Column> SampleTableColumns()
        {
            int i = 1;
            Type objType = typeof(SampleTableModel);
            List<Column> excelColumnDefinitionArray = new()
            {
                //                                            PropertyName                                  format                      col   text
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.Text),                ExcelFormatType.String,     i++), //A
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.Date),                ExcelFormatType.Date,       i++), //B
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.DateAsText),          ExcelFormatType.Date,       i++), //C
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.DateAsGeneral),       ExcelFormatType.Date,       i++), //D
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.Percent),             ExcelFormatType.Percent,    i++), //E
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.PercentAsText),       ExcelFormatType.Percent,    i++), //F
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.PercentAsNumber),     ExcelFormatType.Percent,    i++), //G
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.BoolAsYESNO),         ExcelFormatType.Bool,       i++), //H
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.BoolAsTrueFalse),     ExcelFormatType.Bool,       i++), //I
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.BoolAs10),            ExcelFormatType.Bool,       i++), //J
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.Currency),            ExcelFormatType.Currency,   i++), //K
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.CurrencyAsText),      ExcelFormatType.Currency,   i++), //L
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.CurrencyAsGeneral),   ExcelFormatType.Currency,   i++), //M
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.Decimal),             ExcelFormatType.Decimal,    i++), //N
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.DecimalAsText),       ExcelFormatType.Decimal,    i++), //O
                //TODO ADD DURRATION                                                                                                          
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.RequiredText),        ExcelFormatType.String,     i++), //P
                new Column(_excelExtensionsProvider, objType, nameof(SampleTableModel.OptionalText),        ExcelFormatType.String,     i++), //Q
            };

            return excelColumnDefinitionArray;
        }
    }
}
