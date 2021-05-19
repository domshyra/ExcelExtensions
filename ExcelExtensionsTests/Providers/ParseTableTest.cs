// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using ExcelExtensions.Models;
using ExcelExtensions.Parsers;
using ExcelExtensions.Providers;
using ExcelExtensionsTests.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensionsTests
{
    public class ParseTableTest
    {
        private readonly IExcelExtensionsProvider _excelExtensionsProvider;
        private readonly IParserProvider _parserProvider;
        private readonly TableParser<ParseTableTestBaseModel> _tableParser;

        public ParseTableTest()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _excelExtensionsProvider = new ExcelExtensionsProvider();
            _parserProvider = new ParserProvider();
            _tableParser = new TableParser<ParseTableTestBaseModel>(_excelExtensionsProvider, _parserProvider);
        }

        #region Setup

        [Fact]
        public void Dispose()
        {
            //mockRepository.VerifyAll();
        }

        private static void SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet)
        {
            excel = new ExcelPackage();
            excel.Workbook.Worksheets.Add("Worksheet1");
            worksheet = excel.Workbook.Worksheets["Worksheet1"];
        }

        private void CreateTestFile(ref ExcelWorksheet sheet, List<ParseTableTestBaseModel> models, int headerRow = 1, int startCol = 1)
        {
            int col = startCol;
            sheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{headerRow}"].Value = nameof(ParseTableTestBaseModel.RequiredText);
            sheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{headerRow}"].Value = nameof(ParseTableTestBaseModel.Date);
            sheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{headerRow}"].Value = nameof(ParseTableTestBaseModel.Decimal);
            sheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{headerRow}"].Value = nameof(ParseTableTestBaseModel.OptionalText);


            int row = headerRow + 1;
            col = startCol;
            foreach (ParseTableTestBaseModel testModel in models)
            {
                sheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{row}"].Value = testModel.RequiredText;
                sheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{row}"].Value = testModel.Date;
                sheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{row}"].Value = testModel.Decimal;
                sheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{row}"].Value = testModel.OptionalText;
                //end of row
                row++;
                col = startCol;
            }

        }

        private List<ImportColumnTemplate> GetScannedColumns()
        {
            Type type = typeof(ParseTableTestBaseModel);
            return new List<ImportColumnTemplate>()
            {
                new ImportColumnTemplate(new Column(_excelExtensionsProvider, type, nameof(ParseTableTestBaseModel.RequiredText),   ExcelFormatType.String),   true),
                new ImportColumnTemplate(new Column(_excelExtensionsProvider, type, nameof(ParseTableTestBaseModel.Date),           ExcelFormatType.Date),     true),
                new ImportColumnTemplate(new Column(_excelExtensionsProvider, type, nameof(ParseTableTestBaseModel.Decimal),        ExcelFormatType.Decimal),  true),
                new ImportColumnTemplate(new Column(_excelExtensionsProvider, type, nameof(ParseTableTestBaseModel.OptionalText),   ExcelFormatType.String),   false),
            };
        }
        /// <summary>
        /// For dynamic col test
        /// </summary>
        /// <param name="colStart"></param>
        /// <returns></returns>
        private List<ImportColumnTemplate> GetKnownColumns(int colStart)
        {
            
            int i = colStart;
            Type type = typeof(ParseTableTestBaseModel);
            return new List<ImportColumnTemplate>()
            {
                new ImportColumnTemplate(new Column(_excelExtensionsProvider, type, nameof(ParseTableTestBaseModel.RequiredText),   ExcelFormatType.String,     i++),   true),
                new ImportColumnTemplate(new Column(_excelExtensionsProvider, type, nameof(ParseTableTestBaseModel.Date),           ExcelFormatType.Date,       i++),   true),
                new ImportColumnTemplate(new Column(_excelExtensionsProvider, type, nameof(ParseTableTestBaseModel.Decimal),        ExcelFormatType.Decimal,    i++),   true),
                new ImportColumnTemplate(new Column(_excelExtensionsProvider, type, nameof(ParseTableTestBaseModel.OptionalText),   ExcelFormatType.String,     i++),   false),
            };
        }

        private static List<ParseTableTestBaseModel> GetOneRow()
        {
            return new List<ParseTableTestBaseModel>()
            {
                new ParseTableTestBaseModel()
                {
                    RequiredText = "Test text",
                    Date = new DateTime(2021,1,1),
                    Decimal = 101.3M,
                    OptionalText = null
                }
            };
        }
        /// <summary>
        /// used for <see cref="ParseRow_WithNullRequiredInMiddleOfData"/>
        /// </summary>
        /// <returns></returns>
        private static List<ParseTableTestBaseModel> GetJaggedRows()
        {
            //used for testing parse row method
            return new List<ParseTableTestBaseModel>()
            {
                new ParseTableTestBaseModel()
                {
                    RequiredText = "First Valid Data",
                    Date = new DateTime(2021,1,1),
                    Decimal = 101.3M,
                    OptionalText = "op txt"
                },
                new ParseTableTestBaseModel()
                {
                    RequiredText = null,
                    Date = null,
                    Decimal = null,
                    OptionalText = null
                },
                new ParseTableTestBaseModel()
                {
                    RequiredText = "Second Valid Data",
                    Date = new DateTime(2021,10,10),
                    Decimal = 10,
                    OptionalText = "txt"
                }
            };
        }
        #endregion

        #region Parse 
        [Fact]
        public void ParseTable_BasicTest_Pass()
        {
            //Arrange
            SetUpBasicExcelTest(out _, out ExcelWorksheet worksheet);

            List<ParseTableTestBaseModel> rowsOfTestData = GetOneRow();

            CreateTestFile(ref worksheet, rowsOfTestData);

            //Act
            var result = _tableParser.ParseTable(GetKnownColumns(1), worksheet).Rows;

            //Assert
            Assert.Equal(rowsOfTestData.First(), result.First().Value);
        }

        [Fact]
        public void ParseTable_BasicTest_NoNumbers_Fail()
        {
            //Arrange
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);


            List<ParseTableTestBaseModel> rowsOfTestData = GetOneRow();

            CreateTestFile(ref worksheet, rowsOfTestData);

            //Act
            var ex = Assert.Throws<ArgumentOutOfRangeException>(() => _tableParser.ParseTable(GetScannedColumns(), worksheet).Rows);

            //Assert
            Assert.Equal("All ColumnTemplate's must have a ColumnNumber. If the ColumnNumber is unknown use the ScanForColumnsAndParseTable method. (Parameter 'columns')", ex.Message);
            Assert.Equal("columns", ex.ParamName);
        }

        #endregion Table

        #region Search and parse Table

        [Fact]
        public void SearchAndParseTable_BasicTest_Pass()
        {
            //Arrange
            SetUpBasicExcelTest(out _, out ExcelWorksheet worksheet);

            List<ParseTableTestBaseModel> rowsOfTestData = GetOneRow();

            CreateTestFile(ref worksheet, rowsOfTestData);

            //Act
            var result = _tableParser.ScanForColumnsAndParseTable(GetScannedColumns(), worksheet).Rows;

            //Assert
            Assert.Equal(rowsOfTestData.First(), result.First().Value);
        }

        [Theory]
        [InlineData(3, 1)]
        [InlineData(100, 1)]
        [InlineData(50, 1)]
        [InlineData(101, 10)]
        [InlineData(1, 101)]
        public void SearchAndParseTable_ShiftedLocations_WithHeaderRowInput_Pass(int row, int col)
        {
            //Arrange
            SetUpBasicExcelTest(out _, out ExcelWorksheet worksheet);

            List<ParseTableTestBaseModel> rowsOfTestData = GetOneRow();

            CreateTestFile(ref worksheet, rowsOfTestData, row, col);

            //Act
            var result = _tableParser.ScanForColumnsAndParseTable(GetScannedColumns(), worksheet, row).Rows;

            //Assert
            Assert.Equal(rowsOfTestData.First(), result.First().Value);
        }

        /// <summary>
        /// We should be able to find these columns no matter where they are, so long as they are under max threshold.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        [Theory]
        [InlineData(3, 1)]
        [InlineData(9, 1)]
        [InlineData(50, 1)]
        [InlineData(99, 1)]
        [InlineData(1, 3)]
        [InlineData(1, 9)]
        [InlineData(1, 50)]
        [InlineData(1, 99)]
        [InlineData(3, 3)]
        [InlineData(9, 9)]
        [InlineData(50, 50)]
        [InlineData(99, 99)]
        public void SearchAndParseTable_ShiftedLocations_WithoutHeaderRowInput_Pass(int row, int col)
        {
            //Arrange
            SetUpBasicExcelTest(out _, out ExcelWorksheet worksheet);

            List<ParseTableTestBaseModel> rowsOfTestData = GetOneRow();

            CreateTestFile(ref worksheet, rowsOfTestData, row, col);

            //Act
            var result = _tableParser.ScanForColumnsAndParseTable(GetScannedColumns(), worksheet).Rows;

            //Assert
            Assert.Equal(rowsOfTestData.First(), result.First().Value);
        }
        /// <summary>
        /// Lets have a header row and still need to scan to the for the header row.
        /// </summary>
        [Fact]
        public void SearchAndParseTable_WithHeaderRowInput_LastRowOfSearch_Pass()
        {
            //Arrange
            SetUpBasicExcelTest(out _, out ExcelWorksheet worksheet);

            List<ParseTableTestBaseModel> rowsOfTestData = GetOneRow();

            CreateTestFile(ref worksheet, rowsOfTestData, 104, 1);

            //Act
            var result = _tableParser.ScanForColumnsAndParseTable(GetScannedColumns(), worksheet, 5).Rows;

            //Assert
            Assert.Equal(rowsOfTestData.First(), result.First().Value);
        }
        /// <summary>
        /// Lets import data even though we have a missing optional column. We should see and error and data
        /// </summary>
        [Fact]
        public void SearchAndParseTable_MissingOptionalColumn_Pass()
        {
            //Arrange
            SetUpBasicExcelTest(out _, out ExcelWorksheet worksheet);


            List<ParseTableTestBaseModel> rowsOfTestData = GetOneRow();
            ParseException errorMessage = new()
            {
                ColumnHeader = nameof(ParseTableTestBaseModel.OptionalText),
                Severity = ParseExceptionSeverity.Warning,
                ExceptionType = ParseExceptionType.OptionalFieldMissing
            };

            int col = 1;
            int headerRow = 1;
            int startCol = 1;
            worksheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{headerRow}"].Value = nameof(ParseTableTestBaseModel.RequiredText);
            worksheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{headerRow}"].Value =nameof(ParseTableTestBaseModel.Date);
            worksheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{headerRow}"].Value = nameof(ParseTableTestBaseModel.Decimal);


            int row = headerRow + 1;
            col = startCol;
            foreach (ParseTableTestBaseModel testModel in rowsOfTestData)
            {
                worksheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{row}"].Value = testModel.RequiredText;
                worksheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{row}"].Value = testModel.Date;
                worksheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{row}"].Value = testModel.Decimal;
                //end of row
                row++;
                col = startCol;
            }

            //Act
            var result = _tableParser.ScanForColumnsAndParseTable(GetScannedColumns(), worksheet);

            //Assert
            Assert.Equal(rowsOfTestData.First(), result.Rows.First().Value);
            Assert.Equal(errorMessage.Severity, result.Errors.First().Value.Severity);
            Assert.Equal(errorMessage.ExceptionType, result.Errors.First().Value.ExceptionType);
            Assert.Equal(errorMessage.ColumnHeader, result.Errors.First().Value.ColumnHeader);
        }


        /// <summary>
        /// Let the max threshold be > 100 and see if we get the right error.
        /// </summary>
        [Fact]
        public void SearchAndParseTable_HeaderGreaterThanMaxSearch_Error()
        {
            //Arrange
            SetUpBasicExcelTest(out _, out ExcelWorksheet worksheet);

            List<ParseTableTestBaseModel> rowsOfTestData = GetOneRow();

            CreateTestFile(ref worksheet, rowsOfTestData, 101, 1);

            //Act
            var result = _tableParser.ScanForColumnsAndParseTable(GetScannedColumns(), worksheet).Errors;
            string errorMsg = "Could not find columns. Please check spelling of header columns and make sure all required columns are in the worksheet.";
            //Assert
            Assert.Equal(errorMsg, result.First().Value.Message);
        }

        [Fact]
        public void SearchAndParseTable_MissingRequiredColumn_Error()
        {
            //Arrange
            SetUpBasicExcelTest(out _, out ExcelWorksheet worksheet);

            List<ParseTableTestBaseModel> rowsOfTestData = GetOneRow();
            ParseException errorMessage = new()
            {
                ColumnHeader = nameof(ParseTableTestBaseModel.RequiredText),
                Severity = ParseExceptionSeverity.Error,
                ExceptionType = ParseExceptionType.RequiredFieldMissing
            };

            int col = 1;
            int headerRow = 1;
            int startCol = 1;
            worksheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{headerRow}"].Value = nameof(ParseTableTestBaseModel.Date);
            worksheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{headerRow}"].Value = nameof(ParseTableTestBaseModel.Decimal);
            worksheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{headerRow}"].Value = nameof(ParseTableTestBaseModel.OptionalText);


            int row = headerRow + 1;
            col = startCol;
            foreach (ParseTableTestBaseModel testModel in rowsOfTestData)
            {
                worksheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{row}"].Value = testModel.Date;
                worksheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{row}"].Value = testModel.Decimal;
                worksheet.Cells[$"{_excelExtensionsProvider.GetColumnLetter(col++)}{row}"].Value = testModel.OptionalText;
                //end of row
                row++;
                col = startCol;
            }

            //Act
            ParseException result = _tableParser.ScanForColumnsAndParseTable(GetScannedColumns(), worksheet).Errors.First().Value;

            //Assert
            Assert.Equal(errorMessage.Severity, result.Severity);
            Assert.Equal(errorMessage.ExceptionType, result.ExceptionType);
            Assert.Equal(errorMessage.ColumnHeader, result.ColumnHeader);
        }

        #endregion

        #region ParseRow
        /// <summary>
        /// Lets see if we can parse two things of data where sandwiched in-between is a blank row
        /// </summary>
        [Fact]
        public void ParseRow_WithNullRequiredInMiddleOfData()
        {
            //Arrange
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            List<ParseTableTestBaseModel> rowsOfTestData = GetJaggedRows();
            ParseException errorMessage = new()
            {
                Severity = ParseExceptionSeverity.Warning,
                ExceptionType = ParseExceptionType.MissingData,
                Message = $"Missing all required data for row 3. Skipping row import."
            };
            CreateTestFile(ref worksheet, rowsOfTestData, 1, 1);

            //Act
            var result = _tableParser.ScanForColumnsAndParseTable(GetScannedColumns(), worksheet);

            //Assert
            Assert.Equal(rowsOfTestData.First(x => x.RequiredText == "First Valid Data"), result.Rows[2]);

            var resultErrorMSG = result.Errors.First().Value;

            Assert.Equal(errorMessage.Severity, resultErrorMSG.Severity);
            Assert.Equal(errorMessage.ExceptionType, resultErrorMSG.ExceptionType);
            Assert.Equal(errorMessage.Message, resultErrorMSG.Message);

            Assert.Equal(rowsOfTestData.First(x => x.RequiredText == "Second Valid Data"), result.Rows[4]);
        }

        #endregion
    }
}
