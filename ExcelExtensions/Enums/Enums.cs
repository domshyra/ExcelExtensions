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
            Duration = 4, //TODO this might be a datetime...
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
    }
}
