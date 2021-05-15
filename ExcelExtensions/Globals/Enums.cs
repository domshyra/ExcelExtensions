namespace ExcelExtensions.Enums
{
    public class Enums
    {
        /// <summary>
        /// Format of data to transcribe to and from excel data into c# objects.
        /// </summary>
        public enum ExcelFormatType
        {
            /// <summary>
            /// String type
            /// </summary>
            String = 0,
            /// <summary>
            /// Decimal type
            /// </summary>
            Decimal = 1,
            /// <summary>
            /// Dollar type ($)
            /// </summary>
            Currency = 2,
            /// <summary>
            /// DateTime type
            /// </summary>
            Date = 3,
            /// <summary>
            /// Time span type
            /// </summary>
            Duration = 4,
            /// <summary>
            /// Percent type (%)
            /// </summary>
            Percent = 5,
            /// <summary>
            /// Whole number / int type
            /// </summary>
            Int = 6,
            /// <summary>
            /// Double type
            /// </summary>
            Double = 7,
            /// <summary>
            /// Boolean type
            /// </summary>
            Bool = 8,
            /// <summary>
            /// List of string objects from excel into a dictionary object in c#.
            /// </summary>
            /// <remarks>Does not work with exports</remarks>
            StringList = 9,
            /// <summary>
            /// List of decimal objects from excel into a dictionary object in c#.
            /// </summary>
            /// <remarks>Does not work with exports</remarks>
            DecimalList = 10,
        }

        /// <summary>
        /// Represents the severity of the parser exception
        /// </summary>
        public enum ParseExceptionSeverity
        {
            /// <summary>
            /// Error condition
            /// </summary>
            Error = 0,
            /// <summary>
            /// Warning condition
            /// </summary>
            Warning = 1
        }

        /// <summary>
        /// Represents the exception type the parser ran into 
        /// </summary>
        public enum ParseExceptionType
        {
            DataTypeExpectedError = 0,
            DuplicateData = 1,
            DuplicateKey = 2,
            Generic = 3,
            InvalidData = 4,
            MaxLength = 5,
            MissingData = 6,
            MissMatchData = 7,
            NoColumnTemplatesFound = 8,
            NoFileFound = 9,
            NoValidDataToSave = 10,
            OptionalFieldMissing = 11,
            RequiredFieldMissing = 12,
            SheetMissingError = 13,
            WrongFilePassword = 14
        }
    }

}
