namespace SpreadsheetWriter.Abstractions.Cell
{
    /// <summary>
    /// The definition of a cell address.
    /// </summary>
    public class RowNumber
    {
        /// <summary>
        /// The letter to state the column identifier.
        /// </summary>
        public string Value { get; }

        private RowNumber(string value)
        {
            Value = value;
        }

        /// <summary>
        /// Creates a <see cref="RowNumber">.
        /// </summary>
        public static RowNumber Create(string value)
        {
            return new RowNumber(value);
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            return Value.ToString();
        }
    }
}
