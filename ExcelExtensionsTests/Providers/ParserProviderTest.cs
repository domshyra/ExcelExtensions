// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Globals;
using ExcelExtensions.Providers.Import.Parse;
using OfficeOpenXml;
using System;
using Xunit;

namespace ExcelExtensionsTests
{
    public class ParserProviderTest
    {
        public ParserProviderTest()
        {

        }

        [Fact]
        public void Dispose()
        {
            //mockRepository.VerifyAll();
        }

        private static Parser CreateParserProvider()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            return new Parser();
        }
        private static void SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet)
        {
            excel = new ExcelPackage();
            excel.Workbook.Worksheets.Add("Worksheet1");
            worksheet = excel.Workbook.Worksheets["Worksheet1"];
        }

        #region ParsePercent
        [Theory]
        [InlineData(1.5)]
        [InlineData("1.5")]
        [InlineData(" 1.5 ")]
        [InlineData("1.5%")]
        [InlineData(" 1.5% ")]
        public void ParsePercent_Pass(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;
            decimal? result = provider.ParsePercent(worksheet.Cells["A1"]);

            //Assert
            Assert.Equal(1.5M, result);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void ParsePercent_EmptyAndNull_Fail(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;

            //Assert
            var ex = Assert.Throws<NullReferenceException>(() => provider.ParsePercent(worksheet.Cells["A1"]));
            Assert.Equal($"{Constants.DefaultCannotString} {Constants.DefaultIsNullOrEmptyString} to {Constants.DecimalExceptionTypeString}.", ex.Message);
        }
        [Theory]
        [InlineData(" ")]
        [InlineData("Test")]
        [InlineData("\r")]
        [InlineData("\t")]
        public void ParsePercent_BadData_Fail(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;

            //Assert
            var ex = Assert.Throws<FormatException>(() => provider.ParsePercent(worksheet.Cells["A1"]));
            Assert.Equal($"{Constants.DefaultCannotString} \"{value}\" to {Constants.DecimalExceptionTypeString}.", ex.Message);
        }

        #endregion

        #region ParseInt
        [Theory]
        [InlineData(1)]
        [InlineData("1.0")]
        [InlineData("1")]
        [InlineData(1.0)]
        [InlineData(" 1 ")]
        public void ParseInt_Pass(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;
            int? result = provider.ParseInt(worksheet.Cells["A1"]);

            //Assert
            Assert.Equal(1, result);
        }


        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void ParseInt_EmptyAndNull_Fail(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;

            //Assert
            var ex = Assert.Throws<NullReferenceException>(() => provider.ParseInt(worksheet.Cells["A1"]));
            Assert.Equal($"{Constants.DefaultCannotString} {Constants.DefaultIsNullOrEmptyString} to an int value.", ex.Message);
        }
        [Theory]
        [InlineData(" ")]
        [InlineData("Test")]
        [InlineData("\r")]
        [InlineData("\t")]
        public void ParseInt_BadData_Fail(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;

            //Assert
            var ex = Assert.Throws<FormatException>(() => provider.ParseInt(worksheet.Cells["A1"]));
            Assert.Equal($"{Constants.DefaultCannotString} \"{value}\" to an int value.", ex.Message);
        }
        #endregion

        #region ParseDecimal
        [Theory]
        [InlineData(1)]
        [InlineData("1.0")]
        [InlineData("1")]
        [InlineData(1.0)]
        [InlineData(" 1 ")]
        public void ParseDecimal_Pass(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;
            decimal? result = provider.ParseDecimal(worksheet.Cells["A1"]);

            //Assert
            Assert.Equal(1.0M, result);
        }
        [Theory]
        [InlineData(1.5)]
        [InlineData("1.5")]
        [InlineData(" 1.5 ")]
        public void ParseDecimal_WithDecimalValues_Pass(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;
            decimal? result = provider.ParseDecimal(worksheet.Cells["A1"]);

            //Assert
            Assert.Equal(1.5M, result);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void ParseDecimal_EmptyAndNull_Fail(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;

            //Assert
            var ex = Assert.Throws<NullReferenceException>(() => provider.ParseDecimal(worksheet.Cells["A1"]));
            Assert.Equal($"{Constants.DefaultCannotString} {Constants.DefaultIsNullOrEmptyString} to {Constants.DecimalExceptionTypeString}.", ex.Message);
        }
        [Theory]
        [InlineData(" ")]
        [InlineData("Test")]
        [InlineData("\r")]
        [InlineData("\t")]
        public void ParseDecimal_BadData_Fail(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;

            //Assert
            var ex = Assert.Throws<FormatException>(() => provider.ParseDecimal(worksheet.Cells["A1"]));
            Assert.Equal($"{Constants.DefaultCannotString} \"{value}\" to {Constants.DecimalExceptionTypeString}.", ex.Message);
        }
        #endregion

        #region ParseDouble
        [Theory]
        [InlineData(1)]
        [InlineData("1.0")]
        [InlineData("1")]
        [InlineData(1.0)]
        [InlineData(" 1 ")]
        public void ParseDouble_Pass(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;
            double? result = provider.ParseDouble(worksheet.Cells["A1"]);

            //Assert
            Assert.Equal(1.0, result);
        }
        [Theory]
        [InlineData(1.5)]
        [InlineData("1.5")]
        [InlineData(" 1.5 ")]
        public void ParseDouble_WithDecimalValues_Pass(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;
            double? result = provider.ParseDouble(worksheet.Cells["A1"]);

            //Assert
            Assert.Equal(1.5, result);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void ParseDouble_EmptyAndNull_Fail(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;

            //Assert
            var ex = Assert.Throws<NullReferenceException>(() => provider.ParseDouble(worksheet.Cells["A1"]));
            Assert.Equal($"{Constants.DefaultCannotString} {Constants.DefaultIsNullOrEmptyString} to {Constants.DoubleExceptionTypeString}.", ex.Message);
        }
        [Theory]
        [InlineData(" ")]
        [InlineData("Test")]
        [InlineData("\r")]
        [InlineData("\t")]
        public void ParseDouble_BadData_Fail(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;

            //Assert
            var ex = Assert.Throws<FormatException>(() => provider.ParseDouble(worksheet.Cells["A1"]));
            Assert.Equal($"{Constants.DefaultCannotString} \"{value}\" to {Constants.DoubleExceptionTypeString}.", ex.Message);
        }
        #endregion

        #region ParseDuration

        //Excel starts all time spans from 1899, 12, 30
        [Theory]
        [InlineData("12/30/1899 12:30:00 AM", 0, 30, 0)]
        [InlineData("12/30/1899 1:00:00 AM", 1, 0, 0)]
        [InlineData("12/30/1899 12:00:30 AM", 0, 0, 30)]
        [InlineData("12/30/1899 12:12:12 PM", 12, 12, 12)]
        public void ParseDuration_Pass(object excelValue, int hours, int min, int sec)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = excelValue;
            TimeSpan? result = provider.ParseTimeSpan(worksheet.Cells["A1"]);

            //Assert
            Assert.Equal(sec, result.Value.Seconds);
            Assert.Equal(min, result.Value.Minutes);
            Assert.Equal(hours, result.Value.Hours);
            Assert.Equal(0, result.Value.Milliseconds);
        }



        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void ParseDuration_EmptyAndNull_Fail(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;

            //Assert
            NullReferenceException ex = Assert.Throws<NullReferenceException>(() => provider.ParseTimeSpan(worksheet.Cells["A1"]));
            Assert.Equal($"{Constants.DefaultCannotString} {Constants.DefaultIsNullOrEmptyString} to {Constants.TimeSpanExceptionTypeString}.", ex.Message);
        }
        [Theory]
        [InlineData(" ")]
        [InlineData("Test")]
        [InlineData("\r")]
        [InlineData("\t")]
        public void ParseDuration_BadData_Fail(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;

            //Assert
            var ex = Assert.Throws<FormatException>(() => provider.ParseTimeSpan(worksheet.Cells["A1"]));
            Assert.Equal($"{Constants.DefaultCannotString} \"{value}\" to {Constants.TimeSpanExceptionTypeString}.", ex.Message);
        }

        #endregion

        #region ParseDate

        [Theory]
        [InlineData("12/12/2012")]
        [InlineData(" 12/12/2012 ")]
        [InlineData(41255)]
        [InlineData("41255")]
        [InlineData(" 41255 ")]
        public void ParseDate_Pass(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;
            DateTime? result = provider.ParseDate(worksheet.Cells["A1"]);

            //Assert
            Assert.Equal(12, result.Value.Day);
            Assert.Equal(12, result.Value.Month);
            Assert.Equal(2012, result.Value.Year);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void ParseDate_EmptyAndNull_Fail(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;

            //Assert
            var ex = Assert.Throws<NullReferenceException>(() => provider.ParseDate(worksheet.Cells["A1"]));
            Assert.Equal($"{Constants.DefaultCannotString} {Constants.DefaultIsNullOrEmptyString} to {Constants.DateTimeExceptionTypeString}.", ex.Message);
        }
        [Theory]
        [InlineData(" ")]
        [InlineData("Test")]
        [InlineData("\r")]
        [InlineData("\t")]
        public void ParseDate_BadData_Fail(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;

            //Assert
            var ex = Assert.Throws<FormatException>(() => provider.ParseDate(worksheet.Cells["A1"]));
            Assert.Equal($"{Constants.DefaultCannotString} \"{value}\" to {Constants.DateTimeExceptionTypeString}.", ex.Message);
        }
        #endregion

        #region ParseBool
        [Theory]
        [InlineData(1)]
        [InlineData(" True ")]
        [InlineData(" TRUE ")]
        [InlineData(" true ")]
        [InlineData("True")]
        [InlineData("TRUE")]
        [InlineData("true")]
        [InlineData(" Yes ")]
        [InlineData(" YES ")]
        [InlineData(" yes ")]
        [InlineData("Yes")]
        [InlineData("YES")]
        [InlineData("yes")]
        [InlineData("1")]
        public void ParseBool_Trues_Pass(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;
            bool? result = provider.ParseBool(worksheet.Cells["A1"]);

            //Assert
            Assert.True(result);
        }
        [Theory]
        [InlineData(0)]
        [InlineData(" False ")]
        [InlineData(" FALSE ")]
        [InlineData(" false ")]
        [InlineData("False")]
        [InlineData("FALSE")]
        [InlineData("false")]
        [InlineData(" No ")]
        [InlineData(" NO ")]
        [InlineData(" no ")]
        [InlineData("No")]
        [InlineData("NO")]
        [InlineData("no")]
        [InlineData("0")]
        public void ParseBool_Falses_Pass(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;
            bool? result = provider.ParseBool(worksheet.Cells["A1"]);

            //Assert
            Assert.False(result);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void ParseBool_EmptyAndNull_Fail(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;

            //Assert
            var ex = Assert.Throws<NullReferenceException>(() => provider.ParseBool(worksheet.Cells["A1"]));
            Assert.Equal($"{Constants.DefaultCannotString} {Constants.DefaultIsNullOrEmptyString} to {Constants.BoolExceptionTypeString}.", ex.Message);
        }
        [Theory]
        [InlineData(" ")]
        [InlineData("Test")]
        [InlineData("\r")]
        [InlineData("\t")]
        public void ParseBool_BadData_Fail(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;

            //Assert
            var ex = Assert.Throws<FormatException>(() => provider.ParseBool(worksheet.Cells["A1"]));
            Assert.Equal($"{Constants.DefaultCannotString} \"{value}\" to {Constants.BoolExceptionTypeString}.", ex.Message);
        }

        #endregion

        #region ParseCurrency
        [Theory]
        [InlineData(1.5)]
        [InlineData("1.5")]
        [InlineData(" 1.5 ")]
        [InlineData("$1.5")]
        [InlineData(" $1.5 ")]
        public void ParseCurrency_Pass(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;
            decimal? result = provider.ParseCurrency(worksheet.Cells["A1"]);

            //Assert
            Assert.Equal(1.5M, result);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void ParseCurrency_EmptyAndNull_Fail(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;

            //Assert
            var ex = Assert.Throws<NullReferenceException>(() => provider.ParseCurrency(worksheet.Cells["A1"]));
            Assert.Equal($"{Constants.DefaultCannotString} {Constants.DefaultIsNullOrEmptyString} to {Constants.DecimalExceptionTypeString}.", ex.Message);
        }
        [Theory]
        [InlineData(" ")]
        [InlineData("Test")]
        [InlineData("\r")]
        [InlineData("\t")]
        public void ParseCurrency_BadData_Fail(object value)
        {
            //Arrange
            Parser provider = CreateParserProvider();
            SetUpBasicExcelTest(out ExcelPackage excel, out ExcelWorksheet worksheet);

            //Act
            worksheet.Cells["A1"].Value = value;

            //Assert
            var ex = Assert.Throws<FormatException>(() => provider.ParseCurrency(worksheet.Cells["A1"]));
            Assert.Equal($"{Constants.DefaultCannotString} \"{value}\" to {Constants.DecimalExceptionTypeString}.", ex.Message);
        }
        #endregion

    }
}
