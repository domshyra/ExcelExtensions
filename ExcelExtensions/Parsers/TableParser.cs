// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using ExcelExtensions.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Parsers
{
    public class TableParser<T> : ITableParser<T> where T : class
    {
        private T _model;
        private readonly List<KeyValuePair<int, ParseException>> _singleRowErrors;
        private readonly IExcelExtensionsProvider _excelExtensions;
        private readonly IParserProvider _parser;
        private readonly List<string> _requiredFieldsColumnLocations;
        private readonly ParsedTable<T> _parseResults;
        private readonly List<ParseException> _requiredFieldMissingMessages;


        public TableParser(IExcelExtensionsProvider excelExtensions, IParserProvider parser)
        {
            _excelExtensions = excelExtensions;
            _parser = parser;
            _model = Activator.CreateInstance<T>();
            _singleRowErrors = new List<KeyValuePair<int, ParseException>>();
            _requiredFieldsColumnLocations = new List<string>();
            _parseResults = new ParsedTable<T>();
            _requiredFieldMissingMessages = new List<ParseException>();
        }
        public ParsedTable<T> ScanForColumnsAndParseTable(List<ImportColumnTemplate> columns, ExcelWorksheet workSheet, int headerRowNumber = 1, int maxScanHeaderRowThreashold = 100)
        {
            if (_excelExtensions is null)
            {
                throw new ArgumentNullException(nameof(IExcelExtensionsProvider), $"Make sure when TableParser is constructed that {nameof(IExcelExtensionsProvider)} gets passed in.");
            }
            int rowScanCount = 0;

            List<ImportColumnWithCellAddress> columnsWithCellAddresses = CheckForMissingColumnNumbersInImportColumnTemplates(columns);

            //Check for each of the expected columns in 
            ScanForHeaderRow(columnsWithCellAddresses, ref workSheet, ref headerRowNumber, maxScanHeaderRowThreashold, ref rowScanCount);

            if (CheckMissingScannedColumns(headerRowNumber, maxScanHeaderRowThreashold, rowScanCount, columns))
            {
                return _parseResults;
            }

            //Parse each row
            return ParseRows(columnsWithCellAddresses, workSheet, headerRowNumber);
        }

        public ParsedTable<T> ParseTable(List<ImportColumnTemplate> columns, ExcelWorksheet workSheet, int headerRowNumber = 1)
        {
            if (_excelExtensions is null)
            {
                throw new ArgumentNullException(nameof(IExcelExtensionsProvider), $"Make sure when TableParser is constructed that {nameof(IExcelExtensionsProvider)} gets passed in.");
            }

            List<ImportColumnWithCellAddress> columnsWithCellAddresses =  CheckForMissingColumnNumbersInImportColumnTemplates(columns);

            if (columnsWithCellAddresses.Any(x => x.IsRequired && x.ColumnNumber <= 0))
            {
                throw new ArgumentOutOfRangeException(nameof(columns), "All ColumnTemplate's must have a ColumnNumber. If the ColumnNumber is unknown use the ScanForColumnsAndParseTable method.");
            }
            //Assign required columns
            foreach (ImportColumnWithCellAddress col in columns.Where(x => x.IsRequired))
            {
                _requiredFieldsColumnLocations.Add(col.Column.ColumnLetter);
            }

            //Parse each row
            return ParseRows(columnsWithCellAddresses, workSheet, headerRowNumber);
        }

        /// <summary>
        /// Parse the rows of data, throw out row if all required fields are missing.
        /// </summary>
        /// <param name="columnsWithCellAddresses"></param>
        /// <param name="workSheet"></param>
        /// <param name="headerRowId"></param>
        /// <returns></returns>
        private ParsedTable<T> ParseRows(List<ImportColumnWithCellAddress> columnsWithCellAddresses, ExcelWorksheet workSheet, int headerRowId)
        {
            for (int rowNumber = headerRowId + 1; rowNumber <= workSheet.Dimension.End.Row; rowNumber++)
            {
                //Clear the single errors object
                _singleRowErrors.Clear();
                _model = Activator.CreateInstance<T>();

                foreach (ImportColumnWithCellAddress coltemplate in columnsWithCellAddresses)
                {
                    if (coltemplate.ColumnNumber < 1 || coltemplate.ColumnNumber > workSheet.Dimension.End.Column)
                    {
                        continue;
                    }

                    ExcelRange cell = workSheet.Cells[rowNumber, coltemplate.ColumnNumber];

                    ParseCell(workSheet, rowNumber, coltemplate, cell);
                }

                IList<int> ErrorsFromRequiredFields = _singleRowErrors.Where(x => _requiredFieldsColumnLocations.Contains(x.Value.ColumnLetter)).Select(x => x.Key).ToList();

                if (_singleRowErrors.Count == 0)
                {
                    _parseResults.Rows.Add(rowNumber, _model);
                }
                //All Required fields were blank or contained errors
                else if (ErrorsFromRequiredFields.Count == _requiredFieldsColumnLocations.Count)
                {
                    //skip adding this row since it is only invalid data.

                    ParseException messageSpecification = new()
                    {
                        ExceptionType = ParseExceptionType.MissingData,
                        Severity = ParseExceptionSeverity.Warning,
                        Message = $"Missing all required data for row {rowNumber}. Skipping row import."
                    };

                    KeyValuePair<int, ParseException> missingRow = new(rowNumber, messageSpecification);

                    _parseResults.Errors.Add(missingRow);
                }
                else
                {
                    //This row had both invalid and valid data.
                    _parseResults.Errors.AddRange(_singleRowErrors);
                    _parseResults.Rows.Add(rowNumber, _model);
                }
            }

            return _parseResults;
        }

        /// <summary>
        /// Get the cell address using the extension from the data the user provided in the import map 
        /// </summary>
        /// <param name="columns"></param>
        private List<ImportColumnWithCellAddress> CheckForMissingColumnNumbersInImportColumnTemplates(List<ImportColumnTemplate> columns)
        {
            List<ImportColumnWithCellAddress> columnsWithCellAddresses = new();
            foreach (ImportColumnTemplate item in columns)
            {
                columnsWithCellAddresses.Add(new ImportColumnWithCellAddress(_excelExtensions, item));
            }
            return columnsWithCellAddresses;
        }


        /// <summary>
        /// Try to find the header row of the imported data. If the fields are not found, keep scanning the next row until the row max threshold is reached. 
        /// </summary>
        /// <param name="columns"></param>
        /// <param name="workSheet"></param>
        /// <param name="headerRowNumber"></param>
        /// <param name="maxScanHeaderRowThreashold"></param>
        /// <param name="rowScanCount"></param>
        /// <remarks>There must be at least one required field.</remarks>
        private void ScanForHeaderRow(List<ImportColumnWithCellAddress> columns, ref ExcelWorksheet workSheet, ref int headerRowNumber, int maxScanHeaderRowThreashold, ref int rowScanCount)
        {
            do
            {
                //Make sure the errors are clear first
                _parseResults.Errors.Clear();
                _requiredFieldsColumnLocations.Clear();
                //reset previous rows number 

                FindColumnNamesAndCheckRequiredColumns(columns, ref workSheet, ref headerRowNumber);

                List<ParseException> requiredErrors = _parseResults.Errors.Where(x => x.Value.ExceptionType != ParseExceptionType.OptionalFieldMissing).Select(x => x.Value).ToList();

                if (_requiredFieldMissingMessages.Any() == false)
                {
                    _requiredFieldMissingMessages.AddRange(requiredErrors);
                }
                else if (_requiredFieldMissingMessages.Count > requiredErrors.Count)
                {
                    _requiredFieldMissingMessages.Clear();
                    _requiredFieldMissingMessages.AddRange(requiredErrors);
                }

                if (_parseResults.Errors.Where(x => x.Value.ExceptionType != ParseExceptionType.OptionalFieldMissing).Any() && rowScanCount != maxScanHeaderRowThreashold)
                {
                    //We don't have any columns. Let's bump the header row by one, and try again 
                    headerRowNumber++;
                    rowScanCount++;
                }

            }
            //While we have no missing required columns and while we are not at max scan
            while (_parseResults.Errors.Where(x => x.Value.ExceptionType != ParseExceptionType.OptionalFieldMissing).Any() && rowScanCount != maxScanHeaderRowThreashold);
        }

        /// <summary>
        /// Check for any missing columns when using the scan for header row method.
        /// </summary>
        /// <param name="headerRowId"></param>
        /// <param name="maxScanHeaderRowThreashold"></param>
        /// <param name="rowScanCount"></param>
        /// <returns></returns>
        private bool CheckMissingScannedColumns(int headerRowId, int maxScanHeaderRowThreashold, int rowScanCount, List<ImportColumnTemplate> columns)
        {

            if (rowScanCount >= maxScanHeaderRowThreashold && _requiredFieldMissingMessages.Count == columns.Where(x => x.IsRequired).Count())
            {
                //searched all, but couldn't find one of more in same row
                _parseResults.Errors.Clear();
                ParseException messageSpecification = new()
                {
                    ExceptionType = ParseExceptionType.Generic,
                    Severity = ParseExceptionSeverity.Error,
                    Message = "Could not find columns. Please check spelling of header columns and make sure all required columns are in the worksheet."
                };
                _parseResults.Errors.Add(new KeyValuePair<int, ParseException>(headerRowId, messageSpecification));

                return true;
            }

            if (_requiredFieldMissingMessages.Count > 0)
            {
                _parseResults.Errors.Clear();
                foreach (ParseException msg in _requiredFieldMissingMessages)
                {
                    _parseResults.Errors.Add(new KeyValuePair<int, ParseException>(0, msg));
                }

                return true;
            }

            //The only errors here should be columns missing since no data has been parsed at this point, error if that is true
            if (_parseResults.Errors.Any(x => x.Value.ExceptionType != ParseExceptionType.OptionalFieldMissing))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Parse through the sheet till the header row is found and assign column numbers to each of the col templates.
        /// </summary>
        /// <param name="columns"></param>
        /// <param name="workSheet"></param>
        /// <param name="headerRowId"></param>
        /// <param name="RequiredFields"></param>
        /// <param name="parseResults"></param>
        private void FindColumnNamesAndCheckRequiredColumns(List<ImportColumnWithCellAddress> columns, ref ExcelWorksheet workSheet, ref int headerRowId)
        {
            foreach (ImportColumnWithCellAddress coltemplate in columns)
            {
                foreach (ExcelRangeBase firstRowCell in workSheet.Cells[headerRowId, workSheet.Dimension.Start.Column, headerRowId, workSheet.Dimension.End.Column])
                {
                    if (coltemplate.ColumnHeaderOptions.Any(x => x.Equals(firstRowCell.Text, StringComparison.OrdinalIgnoreCase)))
                    {
                        coltemplate.ColumnNumber = firstRowCell.Start.Column;
                        continue;
                    }
                }

                if (coltemplate.IsRequired && coltemplate.ColumnNumber == 0)
                {

                    ParseException messageSpecification = new(workSheet.Name, coltemplate.Column)
                    {
                        ExceptionType = ParseExceptionType.RequiredFieldMissing,
                        Severity = ParseExceptionSeverity.Error,

                    };
                    _parseResults.Errors.Add(new KeyValuePair<int, ParseException>(headerRowId, messageSpecification));
                }
                else if (coltemplate.ColumnNumber == 0 && coltemplate.IsRequired == false)
                {
                    ParseException messageSpecification = new(workSheet.Name, coltemplate.Column)
                    {
                        ExceptionType = ParseExceptionType.OptionalFieldMissing,
                        Severity = ParseExceptionSeverity.Warning,
                    };

                    _parseResults.Errors.Add(new KeyValuePair<int, ParseException>(headerRowId, messageSpecification));

                }

                if (coltemplate.IsRequired == true)
                {
                    _requiredFieldsColumnLocations.Add(_excelExtensions.GetColumnLetter(coltemplate.ColumnNumber));
                }

            }
        }

        /// <summary>
        /// Try to parse the cell by the columns format type. 
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="rowNumber"></param>
        /// <param name="coltemplate"></param>
        /// <param name="cell"></param>
        /// <remarks>Will throw <see cref="NullReferenceException"/> errors. disable them when debugging through this dll file. </remarks>
        private void ParseCell(ExcelWorksheet workSheet, int rowNumber, ImportColumnTemplate coltemplate, ExcelRange cell)
        {
            switch (coltemplate.Column.Format)
            {
                case ExcelFormatType.Bool:
                    {
                        try
                        {
                            //try to parse
                            bool value = _parser.ParseBool(cell);

                            SetValue(workSheet, rowNumber, coltemplate, cell, value);
                        }
                        catch (NullReferenceException)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogNullReferenceException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        catch (Exception)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogCellException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        break;
                    }
                case ExcelFormatType.Currency:
                    {
                        try
                        {
                            decimal? value = _parser.ParseCurrency(cell);

                            SetValue(workSheet, rowNumber, coltemplate, cell, value);
                        }
                        catch (NullReferenceException)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogNullReferenceException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        catch (Exception)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogCellException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        break;
                    }
                case ExcelFormatType.Date:
                    {
                        try
                        {
                            DateTime? date = _parser.ParseDate(cell);
                            SetValue(workSheet, rowNumber, coltemplate, cell, date);
                            
                        }
                        catch (NullReferenceException)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogNullReferenceException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        catch (Exception)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogCellException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        break;
                    }

                case ExcelFormatType.Duration:
                    {
                        try
                        {
                            TimeSpan? duration = _parser.ParseTimeSpan(cell);
                            SetValue(workSheet, rowNumber, coltemplate, cell, duration);
                        }
                        catch (NullReferenceException)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogNullReferenceException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        catch (Exception)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogCellException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        break;
                    }
                case ExcelFormatType.Percent:
                    {
                        try
                        {
                            decimal? value = _parser.ParsePercent(cell);
                            SetValue(workSheet, rowNumber, coltemplate, cell, value);
                        }
                        catch (NullReferenceException)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogNullReferenceException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        catch (Exception)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogCellException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        break;
                    }
                case ExcelFormatType.Double:
                    {
                        try
                        {
                            double? value = _parser.ParseDouble(cell);
                            SetValue(workSheet, rowNumber, coltemplate, cell, value);
                        }
                        catch (NullReferenceException)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogNullReferenceException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        catch (Exception)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogCellException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        break;
                    }
                case ExcelFormatType.Decimal:
                    {
                        try
                        {
                            decimal? value = _parser.ParseDecimal(cell);
                            SetValue(workSheet, rowNumber, coltemplate, cell, value);
                        }
                        catch (NullReferenceException)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogNullReferenceException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        catch (Exception)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogCellException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        break;
                    }
                case ExcelFormatType.DecimalList:
                    {
                        try
                        {
                            decimal? val = null;
                            try
                            {
                                val = _parser.ParseDecimal(cell);
                            }
                            catch (NullReferenceException)
                            {
                                //Fine if this is null. Lets keep going
                            }


                            // get info about property
                            PropertyInfo modelPropertyInfo = _model.GetType().GetProperty(coltemplate.Column.ModelProperty);

                            var list = modelPropertyInfo.GetValue(_model);
                            Dictionary<string, decimal?> decimalList;
                            if (list == null)  //this is the first item in the list so we need to init the list. 
                            {
                                decimalList = new Dictionary<string, decimal?>();
                            }
                            else
                            {
                                decimalList = (Dictionary<string, decimal?>)list;
                            }

                            decimalList.Add(coltemplate.ColumnHeaderOptions[0], val);

                            modelPropertyInfo.SetValue(_model, decimalList, null);
                        }
                        catch (NullReferenceException)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogNullReferenceException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        catch (Exception)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogCellException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        break;
                    }
                case ExcelFormatType.Int:
                    {
                        try
                        {
                            int? value = _parser.ParseInt(cell);
                            SetValue(workSheet, rowNumber, coltemplate, cell, value);
                        }
                        catch (NullReferenceException)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogNullReferenceException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        catch (Exception)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogCellException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        break;
                    }
                case ExcelFormatType.StringList:
                    {
                        try
                        {
                            _parser.CheckIfCellIsNullOrEmpty(cell, "a string list");

                            string cellVal = cell.Value.ToString().Trim();

                            if (string.IsNullOrEmpty(cellVal) && !string.IsNullOrEmpty(cell.Formula))
                            {
                                //We have a formula which is setting the field to blank instead of null
                                throw new NullReferenceException();
                            }

                            // get info about property
                            PropertyInfo modelPropertyInfo = _model.GetType().GetProperty(coltemplate.Column.ModelProperty);

                            object list = modelPropertyInfo.GetValue(_model);
                            Dictionary<string, string> stringlist;
                            if (list == null)  //this is the first item in the list so we need to init the list. 
                            {
                                stringlist = new Dictionary<string, string>();
                            }
                            else
                            {
                                stringlist = (Dictionary<string, string>)list;
                            }

                            stringlist.Add(coltemplate.ColumnHeaderOptions[0], cellVal);

                            modelPropertyInfo.SetValue(_model, stringlist, null);
                        }
                        catch (NullReferenceException)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogNullReferenceException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        catch (Exception)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogCellException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        break;
                    }
                case ExcelFormatType.String:
                    //uses default
                default:
                    {
                        try
                        {
                            string? value = _parser.ParseString(cell);
                            SetValue(workSheet, rowNumber, coltemplate, cell, value);
                        }
                        catch (NullReferenceException)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogNullReferenceException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        catch (Exception)
                        {
                            _singleRowErrors.Add(_excelExtensions.LogCellException(workSheet.Name, coltemplate, cell.Address, rowNumber));
                        }
                        break;
                    }
            }
        }

        /// <summary>
        /// Try's to set the excel value to the model property
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="rowNumber"></param>
        /// <param name="coltemplate"></param>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        private void SetValue(ExcelWorksheet workSheet, int rowNumber, ImportColumnTemplate coltemplate, ExcelRange cell, object? value)
        {
            try
            {
                // get info about property
                PropertyInfo modelPropertyInfo = _model.GetType().GetProperty(coltemplate.Column.ModelProperty);

                // set value of property
                modelPropertyInfo.SetValue(_model, value, null);
            }
            catch (Exception ex)
            {
                _singleRowErrors.Add(_excelExtensions.LogDeveloperException(workSheet.Name, coltemplate, cell.Address, rowNumber, ex.Message));
            }
        }
    }

}
