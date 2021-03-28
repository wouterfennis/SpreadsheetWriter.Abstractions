
namespace SpreadsheetWriter.Abstractions.Cell
{
    /// <summary>
    /// Abstraction for an range of cells.
    /// </summary>
    public interface ICellRange
    {
        /// <summary>
        /// The address of the cell range.
        /// </summary>
        ICellAddress Address { get; }

        /// <summary>
        /// The value of the cellRange
        /// </summary>
        string Value { get; }
    }
}