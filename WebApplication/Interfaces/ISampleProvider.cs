// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System.Collections.Generic;
using WebApplication.Models;

namespace WebApplication.Interfaces
{
    /// <summary>
    /// Provides methods to use <see cref="ExcelExtensions"/>
    /// </summary>
    public interface ISampleProvider
    {
        /// <summary>
        /// Provides an import of a table with rows of <see cref="SampleTableModel"/> models
        /// </summary>
        /// <param name="file"></param>
        /// <param name="model"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        Dictionary<int, SampleTableModel> ScanForColumnsAndParseTable(IFormFile file, ImportViewModel model, string password);
        /// <summary>
        /// Provides an export of a table with rows of <see cref="SampleTableModel"/> models and fake data
        /// </summary>
        /// <returns></returns>
        ExcelPackage ExportSampleTable();
    }
}