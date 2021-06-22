// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using ExcelExtensions.Models;
using ExcelExtensions.Models.Columns;
using System;
using System.Collections.Generic;
using System.Reflection;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Providers.Import
{
    /// <summary>
    /// Provides list of columns from a <see cref="Type"/> with or without <see cref="ExcelExtensionsColumnAttribute"/>s
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
        /// <para>If no <see cref="FormatType"/> is provided for a property in the <see cref="ExcelExtensionsColumnAttribute"/>s then the <paramref name="formatType"/> will be used.</para>
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
                GetColumnAttributes(modelType, item, formatType, out string modelPropertyName, out string displayName, out ExcelExtensionsColumnAttribute attribute, out FormatType format, out _);

                int columnNumber = GetAttributeColumnNumber(currentColumnNumber, columnsUsed, modelPropertyName, attribute);

                excelColumnDefinitionArray.Add(
                    new ExportColumn(modelPropertyName,
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
        /// <para>If no <see cref="FormatType"/> is provided for a property in the <see cref="ExcelExtensionsColumnAttribute"/>s then the <paramref name="formatType"/> will be used.</para>
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
                GetColumnAttributes(modelType, item, formatType, out string modelPropertyName, out string displayName, out ExcelExtensionsColumnAttribute attribute, out FormatType format, out bool required);

                //TODO: theses all below need to be re though with a new idea of separating the column types out 
                int columnNumber = GetAttributeColumnNumber(currentColumnNumber, columnsUsed, modelPropertyName, attribute);

                excelColumnDefinitionArray.Add(
                    new InformedImportColumn(modelPropertyName,
                                        displayName,
                                        columnNumber,
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
        /// 
        /// </summary>
        /// <param name="currentColumnNumber"></param>
        /// <param name="columnsUsed"></param>
        /// <param name="modelPropertyName"></param>
        /// <param name="attribute"></param>
        /// <returns></returns>
        public int GetAttributeColumnNumber(int currentColumnNumber, Dictionary<int, string> columnsUsed, string modelPropertyName, ExcelExtensionsColumnAttribute attribute)
        {
            bool letterValid = string.IsNullOrEmpty(attribute?.ImportColumnLetter) == false;
            bool numberValid = attribute?.ImportColumnNumber != null;

            int columnNumber;
            if (letterValid && numberValid)
            {
                if (attribute.ImportColumnNumber != _excelExtensions.GetColumnNumber(attribute.ImportColumnLetter))
                {
                    throw new Exception($"Two difference values found for import location both {attribute.ImportColumnNumber} as a number, and {attribute.ImportColumnLetter} as a letter.");
                }
                columnNumber = _excelExtensions.GetColumnNumber(attribute.ImportColumnLetter);
            }
            else if (letterValid)
            {
                columnNumber = _excelExtensions.GetColumnNumber(attribute.ImportColumnLetter);
            }
            else if (numberValid)
            {
                columnNumber = (int)attribute.ImportColumnNumber;
            }
            else
            {
                if (columnsUsed.ContainsKey(currentColumnNumber))
                {
                    throw new Exception($"{columnsUsed[currentColumnNumber]} already uses col number:{currentColumnNumber}(letter:{_excelExtensions.GetColumnLetter(currentColumnNumber)}). Tried to add {currentColumnNumber} since nothing was assigned to {modelPropertyName}");
                }
                else
                {
                    columnsUsed.Add(currentColumnNumber, modelPropertyName);
                }
                columnNumber = currentColumnNumber;
            }

            return columnNumber;
        }

        //todo: unit test
        //1st sets up just for imports in case you need to make tweaks (would have to know the letters or numbers)
        /// <summary>
        /// 
        /// <para>If no <see cref="FormatType"/> is provided for a property in the <see cref="ExcelExtensionsColumnAttribute"/>s then the <paramref name="formatType"/> will be used.</para>
        /// </summary>
        /// <param name="modelType"></param>
        /// <param name="formatType"></param>
        /// <param name="startColumnNumber"></param>
        /// <returns></returns>
        public List<UninformedImportColumn> GetUninformedImportColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1)
        {
            List<UninformedImportColumn> excelColumnDefinitionArray = new();

            foreach (PropertyInfo item in modelType.GetProperties())
            {
                GetColumnAttributes(modelType, item, formatType, out string modelPropertyName, out string displayName, out ExcelExtensionsColumnAttribute attribute, out FormatType format, out bool required);

                //lets make some title options for default if we don't have any
                List<string> titleOptions = attribute?.ImportColumnTitleOptions ?? new List<string>() { modelPropertyName, displayName };

                excelColumnDefinitionArray.Add(
                    new UninformedImportColumn(modelPropertyName,
                                        displayName,
                                        format,
                                        required,
                                        titleOptions,
                                        attribute?.DecimalPrecision));
            }

            return excelColumnDefinitionArray;
        }
        //todo: unit test

        //3rd, option where you really don't care where you export data because it is in order how you set up your model
        /// <summary>
        /// 
        /// <para>If no <see cref="FormatType"/> is provided for a property in the <see cref="ExcelExtensionsColumnAttribute"/>s then the <paramref name="formatType"/> will be used.</para>
        /// </summary>
        /// <param name="modelType"></param>
        /// <param name="formatType"></param>
        /// <param name="startColumnNumber"></param>
        /// <returns></returns>
        public List<InAndOutColumn> GetInAndOutColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1)
        {
            List<InAndOutColumn> excelColumnDefinitionArray = new();
            int currentColumnNumber = startColumnNumber;

            foreach (PropertyInfo item in modelType.GetProperties())
            {
                GetColumnAttributes(modelType, item, formatType, out string modelPropertyName, out string displayName, out ExcelExtensionsColumnAttribute attribute, out FormatType format, out bool required);

                List<string> importKeys = attribute?.ImportColumnTitleOptions ?? new List<string>() { modelPropertyName, displayName };

                excelColumnDefinitionArray.Add(
                    new InAndOutColumn(modelPropertyName,
                                       displayName, //
                                       currentColumnNumber++, //export location
                                       format,
                                       required,
                                       importKeys,
                                       attribute?.DecimalPrecision));

            }
            return excelColumnDefinitionArray;
        }

        //todo: unit test
        public void GetColumnAttributes(Type modelType, PropertyInfo item, FormatType formatType, out string modelPropertyName, out string displayName, out ExcelExtensionsColumnAttribute attribute, out FormatType format, out bool required)
        {
            //read the display name attirube
            modelPropertyName = _excelExtensions.GetExportModelPropertyNameAndDisplayName(modelType, item.Name, out displayName);
            attribute = item.GetCustomAttribute<ExcelExtensionsColumnAttribute>();


            //TODO: throw eroros is tuff is null 

            //Give the default format if not avail be
            format = attribute?.Format ?? formatType;
            required = attribute?.IsRequired ?? true;
        }

    }
}
