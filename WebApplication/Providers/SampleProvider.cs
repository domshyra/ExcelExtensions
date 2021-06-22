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
using ExcelExtensions.Interfaces.Export;
using System.Drawing;
using ExcelExtensions.Models.Import;
using ExcelExtensions.Models.Export;
using ExcelExtensions.Models.Columns.Import;
using ExcelExtensions.Models.Columns;

namespace WebApplication.Providers
{
    public class SampleProvider : ISampleProvider
    {
        private readonly IExtensions _excelExtensions;
        private readonly IFileAndSheetValidatior _validatior;
        private readonly IExporter _exporter;
        private readonly Style _style;

        private readonly string _sheetName = Constants.Constants.SheetName;

        private readonly TableParser<SampleTableModel> _tableParser;

        public SampleProvider(IExtensions excelUtilityProvider, IFileAndSheetValidatior validationProvider, IParser parserProvider, IExporter exporter)
        {
            _excelExtensions = excelUtilityProvider;
            _validatior = validationProvider;
            _exporter = exporter;
            _style = new Style()
            {
                BackgroundColor = ColorTranslator.FromHtml("#001762")
            };

            _tableParser = new TableParser<SampleTableModel>(_excelExtensions, parserProvider);
        }


        public Dictionary<int, SampleTableModel> ScanForColumnsAndParseTable(IFormFile file, ImportViewModel model, string password)
        {
            try
            {
                ExcelPackage package = _validatior.GetExcelPackage(file, password);

                ExcelWorksheet sheet = _validatior.GetExcelWorksheet(package, _sheetName, out ParseException exception);

                if (sheet == null)
                {
                    //Cannot find sheet throw error to the controller
                    throw new ImportException(exception);
                }

                ParsedTable<SampleTableModel> parsedTable = _tableParser.UninformedParseTable(GetTableColumnTemplates(), sheet);

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

        /// <summary>
        /// Sets the requiredments for import and our model property.  
        /// </summary>
        /// <returns></returns>
        private List<UninformedImportColumn> GetTableColumnTemplates()
        {
            //TODO USE REFLECTION https://github.com/domshyra/ExcelExtensions/issues/15
            List<Column> cols = SampleTableColumns();

            return new List<UninformedImportColumn>()
            {
              //                       Col                                                             Req?   
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.Text)),                 true),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.Date)),                 true),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.DateAsText)),           true),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.DateAsGeneral)),        false),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.Duration)),             false),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.Percent)),              true),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.PercentAsText)),        false),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.PercentAsNumber)),      false),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.BoolAsYESNO)),          false),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.BoolAsTrueFalse)),      false),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.BoolAs10)),             false),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.Currency)),             false),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.CurrencyAsText)),       false),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.CurrencyAsGeneral)),    false),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.Decimal)),              false),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.DecimalAsText)),        false),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.RequiredText)),         true),
              new UninformedImportColumn(GetColumn(cols, nameof(SampleTableModel.OptionalText)),         false),
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
        /// <summary>
        /// These are the columns used for excel, the format is needed for the parse and the colNumber is needed for the export column letter
        /// </summary>
        /// <returns></returns>
        private List<Column> SampleTableColumns()
        {
            //TODO USE REFLECTION https://github.com/domshyra/ExcelExtensions/issues/16
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


        public ExcelPackage ExportSampleTable()
        {
            ExcelPackage package = new();
            package.Workbook.Worksheets.Add(_sheetName);
            ExcelWorksheet sheet = package.Workbook.Worksheets[_sheetName];

            List<Column> cols = SampleTableColumns();

            List<SampleTableModel> objects = new() {
                new SampleTableModel() {
                    Text = "Test 1",
                    BoolAs10 = true,
                    Currency = 10.4M,
                    Date = DateTime.Now,
                    RequiredText = "Required Text",
                    Percent = 0.3M,
                    DateAsText = DateTime.Now,
                    Duration = TimeSpan.FromMinutes(30)
                },
                new SampleTableModel() {
                    Text = "Test 2",
                    BoolAs10 = false,
                    Currency = 12.4M,
                    Date = DateTime.Now.AddDays(1),
                    RequiredText = "Required Text Required Text",
                    Percent = 0.4M,
                    DateAsText = DateTime.Now.AddDays(1),
                    Duration = TimeSpan.FromHours(1)
                }
            };



            _exporter.ExportTable(ref sheet, objects, cols, _style);


            return package;
        }
    }
}
