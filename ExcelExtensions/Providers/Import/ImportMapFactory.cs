// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using ExcelExtensions.Models;
using ExcelExtensions.Models.Columns;
using ExcelExtensions.Models.Columns.Import;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static ExcelExtensions.Enums.Enums;

namespace ExcelExtensions.Providers.Import
{
    public class ImportMapFactory
    {
        private readonly IExtensions _excelExtensions;

        public ImportMapFactory(IExtensions excelExtensions)
        {
            _excelExtensions = excelExtensions;

        }

        //If you give me an int you know where it is, or if you give me a letter you know where it is

        //make methods

        //1st sets up just for imports in case you need to make tweaks (would have to know the letters or numbers)


        //2nd set up just for exports (would have to know the letters or numbers)

        //3rd, option where you really don't care where you export data because it is in order how you set up your model


        //Could possible make Column and protected class and make ImportColumn and ExportColum  then a class that encapsulates both for a LazyColumn

        public List<ExportColumn> GetExportColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1)
        {
            List<ExportColumn> excelColumnDefinitionArray = new();

            foreach (PropertyInfo item in modelType.GetProperties())
            {
                GetColumnAttributes(modelType, item, formatType, out string modelPropertyName, out string displayName, out ExcelExtensionsColumnAttribute attribute, out FormatType format, out bool required);

            }
            return excelColumnDefinitionArray;
        }
        public List<InformedImportColumn> GetInformedImportColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1)
        {
            int currentColumnNumber = startColumnNumber;
            List<InformedImportColumn> excelColumnDefinitionArray = new();

            foreach (PropertyInfo item in modelType.GetProperties())
            {
                GetColumnAttributes(modelType, item, formatType, out string modelPropertyName, out string displayName, out ExcelExtensionsColumnAttribute attribute, out FormatType format, out bool required);

                //TODO: theses all below need to be re though with a new idea of separating the column types out 
                string letter = attribute?.ExportColumnLetter ?? _excelExtensions.GetColumnLetter(currentColumnNumber);

                List<string> importKeys = attribute?.ImportColumnTitleOptions ?? new List<string>() { modelPropertyName, displayName };


                //next column
                currentColumnNumber++;
            }

            return excelColumnDefinitionArray;
        }
        public List<UninformedImportColumn> GetUninformedImportColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1)
        {
            List<UninformedImportColumn> excelColumnDefinitionArray = new();

            foreach (PropertyInfo item in modelType.GetProperties())
            {
                GetColumnAttributes(modelType, item, formatType, out string modelPropertyName, out string displayName, out ExcelExtensionsColumnAttribute attribute, out FormatType format, out bool required);

                

                 //<see cref="ParseExceptionSeverity.Error"/>
                List<string> importKeys = attribute?.ImportColumnTitleOptions ?? new List<string>() { modelPropertyName, displayName };

                excelColumnDefinitionArray.Add(
                    new UninformedImportColumn(modelPropertyName,
                                        displayName,
                                        format,
                                        required,
                                        importKeys,
                                        attribute?.DecimalPrecision));
            }

            return excelColumnDefinitionArray;
        }

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

        private void GetColumnAttributes(Type modelType, PropertyInfo item, FormatType formatType, out string modelPropertyName, out string displayName, out ExcelExtensionsColumnAttribute attribute, out FormatType format, out bool required)
        {
            //read the display name attirube
            modelPropertyName = _excelExtensions.GetExportModelPropertyNameAndDisplayName(modelType, item.Name, out displayName);
            attribute = item.GetCustomAttribute<ExcelExtensionsColumnAttribute>();
            //Give the default format if not avail be
            format = attribute?.Format ?? formatType;
            required = attribute?.IsRequired ?? true;
        }

    }
}
