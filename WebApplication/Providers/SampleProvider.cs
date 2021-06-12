// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces.Import;
using ExcelExtensions.Interfaces.Import.Parse;
using ExcelExtensions.Models;
using ExcelExtensions.Providers.Import.Parse;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using WebApplication.Interfaces;
using WebApplication.Models;
using static ExcelExtensions.Enums.Enums;
using ExcelExtensions.Interfaces;

namespace WebApplication.Providers
{
    public class SampleProvider : ISampleProvider
    {
        private readonly IExtensions _excelExtensions;
        private readonly IFileAndSheetValidatior _excelValidatior;

        private readonly TableParser<SampleTableModel> _tableParser;

        public SampleProvider(IExtensions excelUtilityProvider, IFileAndSheetValidatior validationProvider, IParser parserProvider)
        {
            _excelExtensions = excelUtilityProvider;
            _excelValidatior = validationProvider;

            _tableParser = new TableParser<SampleTableModel>(_excelExtensions, parserProvider);
        }


        public Dictionary<int, SampleTableModel> ScanForColumnsAndParseTable(IFormFile file, ImportViewModel model, string password)
        {
            try
            {
                ExcelPackage package = _excelValidatior.GetExcelPackage(file, password);

                ExcelWorksheet sheet = _excelValidatior.GetExcelWorksheet(package, "Sheet1", out ParseException exception);

                if (sheet == null)
                {
                    //Cannot find sheet throw error to the controller
                    throw new ImportException(exception);
                }

                ParsedTable<SampleTableModel> parsedTable = _tableParser.ScanForColumnsAndParseTable(GetTableColumnTemplates(), sheet);

                if (parsedTable.Exceptions.Any(x => x.Value.Severity == ParseExceptionSeverity.Error))
                {
                    //If we encountered any errors in the parse lets throw them to the controller
                    throw new ImportException(parsedTable.Exceptions);
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
            List<Column> cols = SampleTableColumns();

            return new List<ImportColumnTemplate>()
            {
              //                       Col                                                             Req?   
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.Text)),                 true),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.Date)),                 true),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.DateAsText)),           true),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.DateAsGeneral)),        false),
              new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.Duration)),             false),
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
              //new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.ListOfStrings)),        false),
              //new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.ListOfStrings)),        false),
              //new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.ListOfStrings)),        false),
              //new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.ListOfDecimals)),       false),
              //new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.ListOfDecimals)),       false),
              //new ImportColumnTemplate(GetColumn(cols, nameof(SampleTableModel.ListOfDecimals)),       false),
            };
        }

        private static Column GetColumn(List<Column> cols, string modelPropertyName)
        {
            return cols.First(x => x.ModelProperty == modelPropertyName);
        }
        private List<Column> SampleTableColumns()
        {
            int colNumber = 1;
            Type objType = typeof(SampleTableModel);
            List<Column> excelColumnDefinitionArray = new()
            {
                //                                            PropertyName                          format                      col           text
                new Column(_excelExtensions, objType, nameof(SampleTableModel.Text),                FormatType.String,     colNumber++), //A
                new Column(_excelExtensions, objType, nameof(SampleTableModel.Date),                FormatType.Date,       colNumber++), //B
                new Column(_excelExtensions, objType, nameof(SampleTableModel.DateAsText),          FormatType.Date,       colNumber++), //C
                new Column(_excelExtensions, objType, nameof(SampleTableModel.DateAsGeneral),       FormatType.Date,       colNumber++), //D
                new Column(_excelExtensions, objType, nameof(SampleTableModel.Percent),             FormatType.Percent,    colNumber++), //E
                new Column(_excelExtensions, objType, nameof(SampleTableModel.PercentAsText),       FormatType.Percent,    colNumber++), //F
                new Column(_excelExtensions, objType, nameof(SampleTableModel.PercentAsNumber),     FormatType.Percent,    colNumber++), //G
                new Column(_excelExtensions, objType, nameof(SampleTableModel.BoolAsYESNO),         FormatType.Bool,       colNumber++), //H
                new Column(_excelExtensions, objType, nameof(SampleTableModel.BoolAsTrueFalse),     FormatType.Bool,       colNumber++), //I
                new Column(_excelExtensions, objType, nameof(SampleTableModel.BoolAs10),            FormatType.Bool,       colNumber++), //J
                new Column(_excelExtensions, objType, nameof(SampleTableModel.Currency),            FormatType.Currency,   colNumber++), //K
                new Column(_excelExtensions, objType, nameof(SampleTableModel.CurrencyAsText),      FormatType.Currency,   colNumber++), //L
                new Column(_excelExtensions, objType, nameof(SampleTableModel.CurrencyAsGeneral),   FormatType.Currency,   colNumber++), //M
                new Column(_excelExtensions, objType, nameof(SampleTableModel.Decimal),             FormatType.Decimal,    colNumber++), //N
                new Column(_excelExtensions, objType, nameof(SampleTableModel.DecimalAsText),       FormatType.Decimal,    colNumber++), //O
                new Column(_excelExtensions, objType, nameof(SampleTableModel.Duration),            FormatType.Duration,   colNumber++), //P
                new Column(_excelExtensions, objType, nameof(SampleTableModel.RequiredText),        FormatType.String,     colNumber++), //Q
                new Column(_excelExtensions, objType, nameof(SampleTableModel.OptionalText),        FormatType.String,     colNumber++), //R
            };

            return excelColumnDefinitionArray;
        }
    }
}
