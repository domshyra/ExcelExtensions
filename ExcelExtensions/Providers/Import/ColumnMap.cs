// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using ExcelExtensions.Interfaces.Import;
using ExcelExtensions.Models;
using ExcelExtensions.Models.Columns;
using ExcelExtensions.Models.Columns.Export;
using ExcelExtensions.Models.Columns.Import;
using System;
using System.Collections.Generic;
using System.Reflection;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Providers.Import
{
    /// <summary>
    /// Provides list of columns from a <see cref="Type"/> with or without <see cref="ColumnAttribute"/>s
    /// </summary>
    public class ColumnMap : IColumnMap
    {
        private readonly IExtensions _excelExtensions;

        public ColumnMap(IExtensions excelExtensions)
        {
            _excelExtensions = excelExtensions;

        }


        //todo: unit test
        //2nd set up just for exports (would have to know the letters or numbers)
        /// <summary>
        /// <para>If no <see cref="FormatType"/> is provided for a property in the <see cref="ColumnAttribute"/>s then the <paramref name="formatType"/> will be used.</para>
        /// </summary>
        /// <param name="modelType"></param>
        /// <param name="formatType"></param>
        /// <param name="startColumnNumber"></param>
        /// <returns></returns>
        public List<ExportColumn> GetExportColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1)
        {
            int currentColumnNumber = startColumnNumber;
            List<ExportColumn> excelColumnDefinitionArray = new();
            Dictionary<int, string> columnsUsed = new();

            foreach (PropertyInfo item in modelType.GetProperties())
            {
                GetColumnAttributes(modelType, item, formatType, out string modelPropertyName, out string displayName, out ColumnAttribute attribute, out FormatType format, out _);

                int columnNumber = GetAttributeColumnNumber(false, //export
                    currentColumnNumber, columnsUsed, modelPropertyName, attribute);

                excelColumnDefinitionArray.Add(new ExportColumn(modelPropertyName,
                                     displayName,
                                     columnNumber,
                                     format,
                                     attribute?.DecimalPrecision));

                //next column
                currentColumnNumber++;

            }
            return excelColumnDefinitionArray;
        }



        //todo: unit test
        //1st sets up just for imports in case you need to make tweaks (would have to know the letters or numbers)

        /// <summary>
        /// Numbers or Letters are known here and defined in atteributes. Or will import for column order based on propereties
        /// <para>If no <see cref="FormatType"/> is provided for a property in the <see cref="ColumnAttribute"/>s then the <paramref name="formatType"/> will be used.</para>
        /// </summary>
        /// <param name="modelType"></param>
        /// <param name="formatType"></param>
        /// <param name="startColumnNumber"></param>
        /// <returns></returns>
        public List<InformedImportColumn> GetInformedImportColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1)
        {
            int currentColumnNumber = startColumnNumber;
            List<InformedImportColumn> excelColumnDefinitionArray = new();
            Dictionary<int, string> columnsUsed = new();

            foreach (PropertyInfo item in modelType.GetProperties())
            {
                GetColumnAttributes(modelType, item, formatType, out string modelPropertyName, out string displayName, out ColumnAttribute attribute, out FormatType format, out bool required);

                //TODO: theses all below need to be re though with a new idea of separating the column types out 
                int columnNumber = GetAttributeColumnNumber(true, //import
                    currentColumnNumber, columnsUsed, modelPropertyName, attribute);

                excelColumnDefinitionArray.Add(
                    new InformedImportColumn(modelPropertyName,
                                        displayName, //title
                                        columnNumber, //excel column location number 
                                        format,
                                        required,
                                        attribute?.DecimalPrecision));

                //next column
                currentColumnNumber++;
            }

            return excelColumnDefinitionArray;
        }

        //todo: unit test
        /// <summary>
        /// Gets column number from the <see cref="ColumnAttribute."/>
        /// </summary>
        /// <param name="import"></param>
        /// <param name="currentColumnNumber"></param>
        /// <param name="columnsUsed"></param>
        /// <param name="modelPropertyName"></param>
        /// <param name="attribute"></param>
        /// <returns></returns>
        public int GetAttributeColumnNumber(bool import, int currentColumnNumber, Dictionary<int, string> columnsUsed, string modelPropertyName, ColumnAttribute attribute)
        {
            bool letterValid = import ? string.IsNullOrEmpty(attribute?.ImportColumnLetter) == false : string.IsNullOrEmpty(attribute?.ExportColumnLetter) == false;
            bool numberValid = import ? attribute?.ImportColumnNumber != null : attribute?.ExportColumnNumber != null;

            int columnNumber;
            if (letterValid && numberValid)
            {
                int columnLetterValue = getColumnNumberFromLetter(import, attribute);
                int colNum = getColumnNumber(import, attribute);
                
                if (colNum != columnLetterValue)
                {
                    throw new Exception($"Two difference values found for import {ErrorLocation(columnLetterValue)} as letter and {ErrorLocation(colNum)} as number.");
                }
                columnNumber = columnLetterValue;
            }
            else if (letterValid)
            {
                columnNumber = getColumnNumberFromLetter(import, attribute); ;
            }
            else if (numberValid)
            {
                columnNumber = getColumnNumber(import, attribute);
            }
            else
            {
                if (columnsUsed.ContainsKey(currentColumnNumber))
                {
                    throw new Exception($"{columnsUsed[currentColumnNumber]} already uses column {ErrorLocation(currentColumnNumber)}. Tried to add {currentColumnNumber} since nothing was assigned to {modelPropertyName}");
                }
                else
                {
                    columnsUsed.Add(currentColumnNumber, modelPropertyName);
                }
                columnNumber = currentColumnNumber;
            }

            return columnNumber;

            int getColumnNumberFromLetter(bool import, ColumnAttribute attribute)
            {
                return import ? _excelExtensions.GetColumnNumber(attribute.ImportColumnLetter) : _excelExtensions.GetColumnNumber(attribute.ExportColumnLetter);
            }

            static int getColumnNumber(bool import, ColumnAttribute attribute)
            {
                return import ? (int)attribute.ImportColumnNumber : (int)attribute.ExportColumnNumber;
            }
        }

        private string ErrorLocation(int currentColumnNumber) => $"(number:{currentColumnNumber})(letter:{_excelExtensions.GetColumnLetter(currentColumnNumber)}";

        //todo: unit test
        //1st sets up just for imports in case you need to make tweaks (would have to know the letters or numbers)
        /// <summary>
        /// 
        /// <para>If no <see cref="FormatType"/> is provided for a property in the <see cref="ColumnAttribute"/>s then the <paramref name="formatType"/> will be used.</para>
        /// </summary>
        /// <param name="modelType"></param>
        /// <param name="formatType"></param>
        /// <param name="startColumnNumber"></param>
        /// <returns></returns>
        public List<UninformedImportColumn> GetUninformedImportColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1)
        {
            List<UninformedImportColumn> uninformedImportColumns = new();

            foreach (PropertyInfo item in modelType.GetProperties())
            {
                GetColumnAttributes(modelType, item, formatType, out string modelPropertyName, out string displayName, out ColumnAttribute attribute, out FormatType format, out bool required);

                List<string> defualtTitles = new() { modelPropertyName, displayName };
                //if null use default tiles
                List<string> titleOptions = attribute?.ImportColumnTitleOptions ?? defualtTitles;

                //if not null but empty still use the default tiles
                if (titleOptions.Count > 0)
                {
                    titleOptions.AddRange(defualtTitles);
                }

                uninformedImportColumns.Add(new UninformedImportColumn(modelPropertyName,
                                                                       displayName,//title
                                                                       format, 
                                                                       required, 
                                                                       titleOptions, //used to search for in parse table header
                                                                       attribute?.DecimalPrecision));
            }

            return uninformedImportColumns;
        }
        //todo: unit test

        //3rd, option where you really don't care where you export data because it is in order how you set up your model
        /// <summary>
        /// 
        /// <para>If no <see cref="FormatType"/> is provided for a property in the <see cref="ColumnAttribute"/>s then the <paramref name="formatType"/> will be used.</para>
        /// </summary>
        /// <param name="modelType"></param>
        /// <param name="formatType"></param>
        /// <param name="startColumnNumber"></param>
        /// <returns></returns>
        public List<InAndOutColumn> GetInAndOutColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1)
        {
            List<InAndOutColumn> inAndOutColumns = new();
            int currentColumnNumber = startColumnNumber;

            foreach (PropertyInfo item in modelType.GetProperties())
            {
                GetColumnAttributes(modelType, item, formatType, out string modelPropertyName, out string displayName, out ColumnAttribute attribute, out FormatType format, out bool required);

                List<string> importKeys = attribute?.ImportColumnTitleOptions ?? new List<string>() { modelPropertyName, displayName };

                inAndOutColumns.Add(new InAndOutColumn(modelPropertyName,
                                       displayName, //title
                                       currentColumnNumber++, //import/export location
                                       format,
                                       required,
                                       importKeys,
                                       attribute?.DecimalPrecision));

            }
            return inAndOutColumns;
        }

        //todo: unit test
        public void GetColumnAttributes(Type modelType, PropertyInfo item, FormatType formatType, out string modelPropertyName, out string displayName, out ColumnAttribute attribute, out FormatType format, out bool required)
        {
            //read the display name attirube
            modelPropertyName = _excelExtensions.GetExportModelPropertyNameAndDisplayName(modelType, item.Name, out displayName);
            attribute = item.GetCustomAttribute<ColumnAttribute>();


            //TODO: throw eroros is tuff is null 

            //Give the default format if not avail be
            format = attribute?.Format ?? formatType;
            required = attribute?.IsRequired ?? true;
        }

    }
}
