// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

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
        public TimeSpan Duration { get; set; }
        public decimal Percent { get; set; }
        [Display(Name = "Percent as text")]
        public decimal PercentAsText { get; set; }
        [Display(Name = "Percent as number")]
        public decimal PercentAsNumber { get; set; }
        [Display(Name = "Bool as yes no")]
        public bool BoolAsYESNO { get; set; }
        [Display(Name = "Bool as true false")]
        public bool BoolAsTrueFalse { get; set; }
        [Display(Name = "Bool as binary")]
        public bool BoolAs10 { get; set; }
        [Display(Name = "Budget")]
        public decimal Currency { get; set; }
        [Display(Name = "Currency as text")]
        public decimal CurrencyAsText { get; set; }
        [Display(Name = "Currency as feneral")]
        public decimal CurrencyAsGeneral { get; set; }
        public decimal Decimal { get; set; }
        [Display(Name = "Decimal as text")]
        public decimal DecimalAsText { get; set; }
        [Display(Name = "Required text")]
        public string RequiredText { get; set; }
        [Display(Name = "Optional text")]
        public string OptionalText { get; set; }

        public Dictionary<string, string> ListOfStrings { get; set; }
        public Dictionary<string, decimal?> ListOfDecimals { get; set; }
    }
}
