// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelExtensions.Models.Import
{
    /// <summary>
    /// Represents an exception when using the parse functions of excel extensions
    /// </summary>
    public class ImportException : Exception
    {
        public List<ParseException> Exceptions { get; set; }
        public ImportException() : base()
        {
            Exceptions = new List<ParseException>();
        }
        public ImportException(string message) : base(message) { Exceptions = new List<ParseException>(); }

        public ImportException(ParseException exception, string message = null) : base(message)
        {
            Exceptions = new List<ParseException>
            {
                exception
            };
        }
        public ImportException(List<ParseException> exceptions, string message = null) : base(message)
        {
            Exceptions = exceptions;
        }
        public ImportException(List<KeyValuePair<int, ParseException>> exceptions, string message = null) : base(message)
        {
            Exceptions = exceptions.Select(x => x.Value).ToList();
        }

        public ImportException(string message, Exception inner) : base(message, inner) { }
    }
}
