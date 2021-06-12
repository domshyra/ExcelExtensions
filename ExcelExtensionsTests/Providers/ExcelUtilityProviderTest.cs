// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using OfficeOpenXml;
using Xunit;

namespace ExcelExtensionsTests
{
    public class ExcelUtilityProviderTest
    {
        public ExcelUtilityProviderTest()
        {

        }

        [Fact]
        public void Dispose()
        {
            //mockRepository.VerifyAll();
        }

        private static Extensions.Providers.Extension.Extensions CreateProvider()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            return new Extensions.Providers.Extension.Extensions();
        }

        [Theory]
        [InlineData("A", 1)]
        [InlineData("B", 2)]
        [InlineData("C", 3)]
        [InlineData("J", 10)]
        [InlineData("AA", 27)]
        [InlineData("BA", 53)]
        [InlineData("BZ", 78)]
        [InlineData("SF", 500)]
        [InlineData("ZZ", 702)]
        [InlineData("ALL", 1000)]
        public void GetColumnNumber(string colName, int colNumber)
        {
            Extensions.Providers.Extension.Extensions provider = CreateProvider();


            Assert.Equal(colNumber, provider.GetColumnNumber(colName));
        }

        [Theory]
        [InlineData("A1", 1)]
        [InlineData("B2", 2)]
        [InlineData("C3", 3)]
        [InlineData("J10", 10)]
        [InlineData("AA27", 27)]
        [InlineData("BA53", 53)]
        [InlineData("BZ78", 78)]
        [InlineData("SF500", 500)]
        [InlineData("ZZ702", 702)]
        [InlineData("ALL1000", 1000)]
        public void GetColumnNumberAsCell(string cell, int colNumber)
        {
            Extensions.Providers.Extension.Extensions provider = CreateProvider();


            Assert.Equal(colNumber, provider.GetColumnNumber(cell));
        }

        [Theory]
        [InlineData("A", 1)]
        [InlineData("B", 2)]
        [InlineData("C", 3)]
        [InlineData("J", 10)]
        [InlineData("AA", 27)]
        [InlineData("BA", 53)]
        [InlineData("BZ", 78)]
        [InlineData("SF", 500)]
        [InlineData("ZZ", 702)]
        [InlineData("ALL", 1000)]
        public void GetColumnName(string colName, int colNumber)
        {
            Extensions.Providers.Extension.Extensions provider = CreateProvider();


            Assert.Equal(colName, provider.GetColumnLetter(colNumber));
        }

        [Theory]
        [InlineData("A1", 1)]
        [InlineData("B2", 2)]
        [InlineData("C3", 3)]
        [InlineData("J10", 10)]
        [InlineData("AA27", 27)]
        [InlineData("BA53", 53)]
        [InlineData("BZ78", 78)]
        [InlineData("SF500", 500)]
        [InlineData("ZZ702", 702)]
        [InlineData("ALL1000", 1000)]
        public void GetRowNumber(string cell, int rowNumber)
        {
            Extensions.Providers.Extension.Extensions provider = CreateProvider();


            Assert.Equal(rowNumber, provider.GetRowNumber(cell));
        }


        /// <summary>
        /// Luckily the int is the same col and row number ;)
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowNumberAndColNumber"></param>
        [Theory]
        [InlineData("A1", 1)]
        [InlineData("B2", 2)]
        [InlineData("C3", 3)]
        [InlineData("J10", 10)]
        [InlineData("AA27", 27)]
        [InlineData("BA53", 53)]
        [InlineData("BZ78", 78)]
        [InlineData("SF500", 500)]
        [InlineData("ZZ702", 702)]
        [InlineData("ALL1000", 1000)]
        public void GetCellAddress(string cell, int rowNumberAndColNumber)
        {
            Extensions.Providers.Extension.Extensions provider = CreateProvider();

            Assert.Equal(cell, provider.GetCellAddress(rowNumberAndColNumber, rowNumberAndColNumber));
        }

        /// <summary>
        /// Luckily the int is the same col and row number ;)
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowNumberAndColNumber"></param>
        [Theory]
        [InlineData("A1", 1)]
        [InlineData("B2", 2)]
        [InlineData("C3", 3)]
        [InlineData("J10", 10)]
        [InlineData("AA27", 27)]
        [InlineData("BA53", 53)]
        [InlineData("BZ78", 78)]
        [InlineData("SF500", 500)]
        [InlineData("ZZ702", 702)]
        [InlineData("ALL1000", 1000)]
        public void GetCellAddressWithStringRow(string cell, int rowNumberAndColNumber)
        {
            Extensions.Providers.Extension.Extensions provider = CreateProvider();

            Assert.Equal(cell, provider.GetCellAddress(rowNumberAndColNumber, rowNumberAndColNumber.ToString()));
        }
        /// <summary>
        /// Luckily the int is the same col and row number ;)
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowNumberAndColNumber"></param>
        [Theory]
        [InlineData("A1", 1)]
        [InlineData("B2", 2)]
        [InlineData("C3", 3)]
        [InlineData("J10", 10)]
        [InlineData("AA27", 27)]
        [InlineData("BA53", 53)]
        [InlineData("BZ78", 78)]
        [InlineData("SF500", 500)]
        [InlineData("ZZ702", 702)]
        [InlineData("ALL1000", 1000)]
        public void GetCellAddressTuple(string cell, int rowNumberAndColNumber)
        {
            Extensions.Providers.Extension.Extensions provider = CreateProvider();

            Assert.Equal((rowNumberAndColNumber, rowNumberAndColNumber), provider.GetCellAddress(cell));
        }
    }
}
