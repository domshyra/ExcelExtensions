// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Extensions.Models
{
    public class Style
    {
        public bool Bold { get; internal set; }
        public ExcelFillStyle PatternType { get; internal set; }
        public Color BackgroundColor { get; internal set; }
        public Color Color { get; internal set; }
    }
}
