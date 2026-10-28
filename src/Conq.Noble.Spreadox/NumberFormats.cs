namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// Common Excel number-format codes for <see cref="ExcelCellStyle.NumberFormat"/>.
    /// </summary>
    public static class NumberFormats
    {
        public const string General = "General";
        public const string Integer = "0";
        public const string TwoDecimals = "0.00";
        public const string Thousands = "#,##0";
        public const string ThousandsTwoDecimals = "#,##0.00";
        public const string Percent = "0%";
        public const string PercentTwoDecimals = "0.00%";
        public const string Date = "yyyy-mm-dd";
        public const string DateTime = "yyyy-mm-dd hh:mm:ss";
        public const string Time = "hh:mm:ss";
        public const string Currency = "$#,##0.00";
    }
}
