// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtensions.Models
{
    public class LazyColumn : Column
    {
        public int ColumnLocation { get; set; }

        public LazyColumn()
        {

        }
    }
}
