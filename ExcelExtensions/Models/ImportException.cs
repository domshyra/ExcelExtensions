// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using System;
using System.Collections.Generic;

//TODO
namespace ExcelExtensions.Models
{
    /// <summary>
    /// Represents an exception when using the parse functions of EE
    /// </summary>
    public class ImportException : Exception
    {
        public List<ParseException> Exceptions { get; set; }
        public ImportException() : base()
        {
            Exceptions = new List<ParseException>();
        }
        public ImportException(string message) : base(message) { Exceptions = new List<ParseException>(); }

        public ImportException(ParseException exception) : base()
        {
            Exceptions = new List<ParseException>
            {
                exception
            };
        }
        public ImportException(List<ParseException> exceptions) : base()
        {
            Exceptions = exceptions;
        }

        public ImportException(string message, Exception inner) : base(message, inner) { }
    }
}
