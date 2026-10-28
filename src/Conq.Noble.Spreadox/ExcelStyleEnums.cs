namespace Conq.Noble.Spreadox
{
    public enum ExcelHorizontalAlignment
    {
        General = 0,
        Left,
        Center,
        Right,
        Fill,
        Justify,
        CenterContinuous,
        Distributed
    }

    public enum ExcelVerticalAlignment
    {
        Top = 0,
        Center,
        Bottom,
        Justify,
        Distributed
    }

    public enum ExcelBorderStyle
    {
        None = 0,
        Thin,
        Medium,
        Thick,
        Dashed,
        Dotted,
        Double,
        Hair,
        MediumDashed,
        DashDot,
        MediumDashDot,
        DashDotDot,
        MediumDashDotDot,
        SlantDashDot
    }

    public enum ExcelFillPattern
    {
        None = 0,
        Solid
    }

    public enum ConditionalOperator
    {
        GreaterThan,
        GreaterThanOrEqual,
        LessThan,
        LessThanOrEqual,
        Equal,
        NotEqual
    }
}
