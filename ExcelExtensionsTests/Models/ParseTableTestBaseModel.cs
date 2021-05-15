// Copyright (c) Dominic Schira <domshyra@gmail.com>. All Rights Reserved.

using System;

namespace ExcelExtensionsTests.Models
{
    public class ParseTableTestBaseModel
    {
        public string RequiredText { get; set; }
        public DateTime? Date { get; set; }
        public decimal? Decimal { get; set; }
        public string OptionalText { get; set; }

        public ParseTableTestBaseModel()
        {

        }

        public override int GetHashCode()
        {
            //We don't care about overflow here just need to make a hash
            unchecked
            {
                return 31 * RequiredText.GetHashCode() + Date.GetHashCode() + Decimal.GetHashCode() + OptionalText.GetHashCode();
            }
        }

        /// <summary>
        /// Override default equals since this is an object and the normal one will be a pointer compare.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns>True if both ItemCodeDescriptorId are equivalent</returns>
        public override bool Equals(object obj)
        {
            if (obj is ParseTableTestBaseModel compareModel)
            {
                if (RequiredText == null && compareModel.RequiredText == null)
                {
                    return true;
                }
                else if (RequiredText == null || compareModel.RequiredText == null)
                {
                    return false;
                }
                return  string.Equals(RequiredText.Trim(), compareModel.RequiredText.Trim(), StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(Date?.ToShortDateString(), compareModel.Date?.ToShortDateString(), StringComparison.OrdinalIgnoreCase) &&
                        decimal.Round((decimal)Decimal, 2) == decimal.Round((decimal)compareModel.Decimal, 2) &&
                        string.Equals(OptionalText?.Trim(), compareModel.OptionalText?.Trim(), StringComparison.OrdinalIgnoreCase);
            }
            return base.Equals(obj);
        }
    }
}

