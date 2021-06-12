// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using OfficeOpenXml.Style;
using System.Drawing;

namespace ExcelExtensions.Models
{
    public class Style
    {
        public bool Bold { get; set; }
        public ExcelFillStyle PatternType { get; set; }
        public Color BackgroundColor { get; set; }
        public Color Color { get; set; }

        public Style()
        {

        }
    }
}
