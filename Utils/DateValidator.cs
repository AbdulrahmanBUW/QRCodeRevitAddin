using System;
using System.Globalization;

namespace QRCodeRevitAddin.Utils
{
    public static class DateValidator
    {
        private const string DateFormatStandard = "dd/MM/yy";
        private const string DateFormatAlternate = "dd/MM/yyyy";

        public static bool IsValid(string dateString)
        {
            if (string.IsNullOrWhiteSpace(dateString))
                return false;

            DateTime result;

            if (DateTime.TryParseExact(dateString, DateFormatStandard, CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
                return true;

            if (DateTime.TryParseExact(dateString, DateFormatAlternate, CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
                return true;

            return false;
        }

        public static string GetTodayFormatted()
        {
            return DateTime.Now.ToString(DateFormatStandard);
        }

        public static string Format(DateTime date)
        {
            return date.ToString(DateFormatStandard);
        }

        public static string GetExpectedFormat()
        {
            return "DD/MM/YY or DD/MM/YYYY";
        }
    }
}