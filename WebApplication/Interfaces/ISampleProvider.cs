// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
using WebApplication.Models;

namespace WebApplication.Interfaces
{
    public interface ISampleProvider
    {
        Dictionary<int, SampleTableModel> ScanForColumnsAndParseTable(IFormFile file, ImportViewModel model, string password);
    }
}