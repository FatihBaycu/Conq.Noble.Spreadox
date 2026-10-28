namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// Discriminator for values stored in a cell.
    /// </summary>
    public enum CellValueKind
    {
        Empty = 0,
        String,
        Integer,
        Double,
        Decimal,
        Boolean,
        DateTime,
        /// <summary>Cell stores a formula string (optional cached value in <see cref="ExcelCell.Value"/>).</summary>
        Formula
    }
}
