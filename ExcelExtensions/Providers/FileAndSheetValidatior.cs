// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using ExcelExtensions.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Providers
{
    /// <inheritdoc/>
    public class FileAndSheetValidatior : IFileAndSheetValidatior
    {
        public FileAndSheetValidatior()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <inheritdoc/>
        public ExcelWorksheet GetExcelWorksheet(ExcelPackage package, string sheetName, out ParseException exception)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

            if (worksheet == null)
            {
                //We don't have a valid sheet
                exception = new()
                {
                    Sheet = sheetName,
                    ExceptionType = ParseExceptionType.SheetMissingError,
                    Severity = ParseExceptionSeverity.Error,
                };

                return null;
            }

            exception = null;
            return worksheet;
        }

        /// <inheritdoc/>
        public ExcelPackage GetExcelPackage(dynamic excelFile, string password = "")
        {
            ParseException exception = null;

            //Make sure we have a file
            if (excelFile == null)
            {
                exception = new()
                {
                    ExceptionType = ParseExceptionType.NoFileFound,
                    Severity = ParseExceptionSeverity.Error,
                };
            }

            if (PropertyExist(excelFile, "ContentType") == false || PropertyExist(excelFile, "FileName") == false)
            {
                throw new FormatException("file must be type IFormFile or HttpPostedFileBase");
            }

            //TODO: accept more formats later
            //Make sure the file is an excel type
            if (excelFile.ContentType != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                exception = new()
                {
                    Message = "Wrong file type uploaded. Please upload an excel file",
                    ExceptionType = ParseExceptionType.Generic,
                    Severity = ParseExceptionSeverity.Error,
                };
            }

            if (exception != null)
            {
                throw new ImportException(exception);
            }

            //Core
            if (PropertyExist(excelFile, "InputStream") == false)
            {
                try
                {
                    if (!string.IsNullOrEmpty(password))
                    {
                        //We need to create a read/write memory stream that is needed to read a password protected file
                        MemoryStream readWrite = new();
                        excelFile.CopyTo(readWrite);
                        return new ExcelPackage(readWrite, password);
                    }
                    else
                    {
                        return new ExcelPackage(excelFile.OpenReadStream());
                    }
                }
                catch (InvalidDataException)
                {
                    //the password is not supplied to a password protected file
                    throw new ImportException(PasswordError("not supplied"));
                }
                catch (System.Security.SecurityException)
                {
                    //the password supplied is invalid
                    throw new ImportException(PasswordError("invalid"));
                }
            }

            //MVC
            else
            {
                try
                {
                    if (!string.IsNullOrEmpty(password))
                    {
                        return new ExcelPackage(excelFile.InputStream, password);
                    }
                    else
                    {
                        return new ExcelPackage(excelFile.InputStream);
                    }
                }
                catch (InvalidDataException)
                {
                    throw new ImportException(PasswordError("invalid"));
                }
            }
        }

        private static ParseException PasswordError(string passwordStatus)
        {
            ParseException exception = new()
            {
                Message = $"This file is encrypted with a password, which was {passwordStatus}.",
                ExceptionType = ParseExceptionType.WrongFilePassword,
                Severity = ParseExceptionSeverity.Error,
            };

            return exception;
        }

        private static bool PropertyExist(dynamic excelFile, string propertyName)
        {
            if (excelFile is ExpandoObject)
                return ((IDictionary<string, object>)excelFile).ContainsKey(propertyName);

            return excelFile.GetType().GetProperty(propertyName) != null;
        }

    }
}
