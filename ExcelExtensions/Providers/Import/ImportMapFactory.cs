// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Interfaces;
using ExcelExtensions.Models;
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


        public List<UninformedImportColumn> MapObjectToImportColumns(Type modelType, FormatType formatType = FormatType.String, int startColumnNumber = 1)
        {
            //If you give me an int you know where it is, or if you give me a letter you know where it is

            //make methods

            //1st sets up just for imports in case you need to make tweaks (would have to know the letters or numbers)


            //2nd set up just for exports (would have to know the letters or numbers)

            //3rd, option where you really don't care where you export data because it is in order how you set up your model


            //Could possible make Column and protected class and make ImportColumn and ExportColum  then a class that encapsulates both for a LazyColumn

            int currentColumnNumber = startColumnNumber;
            List<UninformedImportColumn> excelColumnDefinitionArray = new();

            foreach (PropertyInfo item in modelType.GetProperties())
            {
                //read the display name attirube
                string modelPropertyName = _excelExtensions.GetExportModelPropertyNameAndDisplayName(modelType, item.Name, out string textTitle);

                ExcelExtensionsColumnAttribute attribute = item.GetCustomAttribute<ExcelExtensionsColumnAttribute>();

                //Give the default format if not avail be
                FormatType format = attribute?.Format ?? formatType;

                //TODO: theses all below need to be re though with a new idea of separating the column types out 
                string letter = attribute?.ExportColumnLetter ?? _excelExtensions.GetColumnLetter(currentColumnNumber);
                bool required = attribute?.IsRequired ?? true; //<see cref="ParseExceptionSeverity.Error"/>
                List<string> importKeys = attribute?.ImportColumnTitleOptions ?? new List<string>() { modelPropertyName, textTitle };

                excelColumnDefinitionArray.Add(new InformedImportColumn()
                {
                    //If already know our column letter lets use that. 
                    ImportColumnNumber = attribute?.ExportColumnLetter == null ? 0 : _excelExtensions.GetColumnNumber(attribute.ExportColumnLetter),
                    Column = new Column(modelPropertyName,
                                        textTitle,
                                        letter,
                                        format,
                                        attribute?.DecimalPrecision),
                    DisplayNameOptions = importKeys,
                    IsRequired = required
                });
                //next column
                currentColumnNumber++;
            }

            return excelColumnDefinitionArray;
        }
    }
}
