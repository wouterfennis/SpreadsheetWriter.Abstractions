
namespace SpreadsheetWriter.Abstractions.Formula
{
    /// <summary>
    /// Factory for creating <see cref="IFormulaBuilder"/>.
    /// </summary>
    public interface IFormulaBuilderFactory
    {
        /// <summary>
        /// Creates a <see cref="IFormulaBuilder"/>.
        /// </summary>
        IFormulaBuilder Create();
    }
}
