﻿// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using Extensions.Models;
using System.Collections.Generic;

namespace WebApplication.Models
{
    public class ImportViewModel
    {
        public List<ParseException> Exceptions { get; set; }
        public string ScanForColumnsAndParseTable { get; internal set; }

        public ImportViewModel()
        {
            Exceptions = new List<ParseException>();
            ScanForColumnsAndParseTable = "[]";
        }
    }
}
