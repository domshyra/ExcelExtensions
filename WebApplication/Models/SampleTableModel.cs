// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication.Models
{
    public class SampleTableModel
    {
        [Display(Name = "City")]
        public string Text { get; set; }
        [Display(Name = "Date as date")]
        public DateTime Date { get; set; }
        public DateTime DateAsText { get; set; }
        public DateTime DateAsGeneral { get; set; }
        public decimal Percent { get; set; }
        public decimal PercentAsText { get; set; }
        public decimal PercentAsNumber { get; set; }
        public bool BoolAsYESNO { get; set; }
        public bool BoolAsTrueFalse { get; set; }
        public bool BoolAs10 { get; set; }
        [Display(Name = "Budget")]
        public decimal Currency { get; set; }
        public decimal CurrencyAsText { get; set; }
        public decimal CurrencyAsGeneral { get; set; }
        public decimal Decimal { get; set; }
        public decimal DecimalAsText { get; set; }
        [Display(Name = "Location")]
        public string RequiredText { get; set; }
        public string OptionalText { get; set; }

        public Dictionary<string, string> ListOfStrings { get; set; }
        public Dictionary<string, decimal?> ListOfDecimals { get; set; }
    }
}
