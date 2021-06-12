// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using OfficeOpenXml.Style;
using System.Drawing;

namespace ExcelExtensions.Models
{
    public class Style
    {
        /// <summary>
        /// Bold font face
        /// <para>Default is false</para>
        /// </summary>
        public bool Bold { get; set; }
        /// <summary>
        /// PatternType
        /// <para>Default is <see cref="ExcelFillStyle.Solid"/></para>
        /// </summary>
        public ExcelFillStyle PatternType { get; set; }
        /// <summary>
        /// Table title row background color
        /// <para>Default is <see cref="Color.Black"/></para>
        /// </summary>
        public Color BackgroundColor { get; set; }
        /// <summary>
        /// Table title row text color
        /// <para>Default is <see cref="Color.White"/></para>
        /// </summary>
        public Color Color { get; set; }

        public Style()
        {
            BackgroundColor = Color.Black;
            Color = Color.White;

            PatternType = ExcelFillStyle.Solid;
        }
    }
}
