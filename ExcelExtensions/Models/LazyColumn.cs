// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtensions.Models
{
    /// <summary>
    /// Eazy button class 
    /// </summary>
    public class LazyColumn : UninformedImportColumn
    {
        /// <summary>
        /// same as in as out
        /// </summary>
        public int ColumnLocation { get; set; }

        public LazyColumn()
        {

        }
    }
}
