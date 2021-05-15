// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using System;

//TODO
namespace ExcelExtensions.Models
{
    /// <summary>
    /// Represents an exception when using the parse functions of EE
    /// </summary>
    public class ImportException : Exception
    {
        public ParseException Exception { get; set; }
        public ImportException() : base() { }
        public ImportException(string message) : base(message) { }

        public ImportException(ParseException exception) => Exception = exception;

        public ImportException(string message, Exception inner) : base(message, inner) { }
    }
}
