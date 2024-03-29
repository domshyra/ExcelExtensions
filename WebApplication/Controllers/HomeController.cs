﻿// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using ExcelExtensions.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.IO;
using WebApplication.Interfaces;
using WebApplication.Models;

namespace WebApplication.Controllers
{
    public class HomeController : Controller
    {
        private readonly ISampleProvider _sampleProvider;

        public HomeController(ISampleProvider sampleProvider)
        {
            _sampleProvider = sampleProvider;
        }
        public IActionResult Index(ImportViewModel form = null)
        {
            if (form == null)
            {
                form = new ImportViewModel();
            }
            return View("Index", JsonConvert.SerializeObject(form));
        }

        public ActionResult ScanForColumnsAndParseTable(ImportViewModel form, IFormFile file, string password = "")
        {
            try
            {
                form.ScanForColumnsAndParseTable = _sampleProvider.ScanForColumnsAndParseTable(file, form, password);
            }
            catch (ImportException ex)
            {
                form.Exceptions = ex.Exceptions;
            }

            return View("Index", JsonConvert.SerializeObject(form));

        }
        public ActionResult ExportTableSample()
        {
            ExcelPackage package = _sampleProvider.ExportSampleTable();
            string fileName = "Table example";
            return Export(package, ref fileName);
        }

        private ActionResult Export(ExcelPackage package, ref string fileName)
        {
            fileName = string.Join("_", fileName.Split(Path.GetInvalidFileNameChars()));
            Response.Headers["content-disposition"] = $"attachment;  filename={fileName}.xlsx";
            return File(package.GetAsByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }
    }
}
