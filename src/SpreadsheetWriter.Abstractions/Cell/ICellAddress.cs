namespace SpreadsheetWriter.Abstractions.Cell
{
    /// <summary>
    /// Abstraction for a cell address.
    /// </summary>
    public interface ICellAddress
    {
        /// <summary>
        /// The column letter of a cell address.
        /// </summary>
        ColumnLetter ColumnLetter { get; }

        /// <summary>
        /// The row number of a cell address.
        /// </summary>
        RowNumber RowNumber { get; }

        /// <summary>
        /// Creates string representation of a cell address.
        /// </summary>
        string ToString();
    }
}