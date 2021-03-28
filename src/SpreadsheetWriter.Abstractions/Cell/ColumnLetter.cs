namespace SpreadsheetWriter.Abstractions.Cell
{
    /// <summary>
    /// The definition of a cell address.
    /// </summary>
    public class ColumnLetter
    {
        public string Value { get; }

        /// <summary>
        /// Constructor.
        /// </summary>
        private ColumnLetter(string columnLetter)
        {
            Value = columnLetter;
        }

        /// <summary>
        /// Creates a <see cref="ColumnLetter">.
        /// </summary>
        public static ColumnLetter Create(string value)
        {
            return new ColumnLetter(value);
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            return Value.ToString();
        }
    }
}
