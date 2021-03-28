using SpreadsheetWriter.Abstractions.Cell;

namespace SpreadsheetWriter.Abstractions.Formula
{
    /// <summary>
    /// Interface to build a spreadsheet formula fluently.
    /// </summary>
    public interface IFormulaBuilder
    {
        /// <summary>
        /// Add address of an cell to the formula.
        /// </summary>
        IFormulaBuilder AddCellAddress(ICellAddress cellAddress);

        /// <summary>
        /// Add the letter of an column to the formula.
        /// </summary>
        IFormulaBuilder AddCellColumnLetter(ColumnLetter columnLetter);

        /// <summary>
        /// Add row number to the formula.
        /// </summary>
        IFormulaBuilder AddRowNumber(RowNumber rowNumber);

        /// <summary>
        /// Add a closing parenthesis to the formula.
        /// </summary>
        IFormulaBuilder AddClosingParenthesis();

        /// <summary>
        /// Add a division sign to the formula.
        /// </summary>
        IFormulaBuilder AddDivisionSign();

        /// <summary>
        /// Add a multiplication sign to the formula.
        /// </summary>
        IFormulaBuilder AddMultiplicationSign();

        /// <summary>
        /// Add a opening parenthesis to the formula.
        /// </summary>
        IFormulaBuilder AddOpenParenthesis();

        /// <summary>
        /// Add a subtraction sign to the formula.
        /// </summary>
        IFormulaBuilder AddSubtractionSign();

        /// <summary>
        /// Add a summation sign to the formula.
        /// </summary>
        IFormulaBuilder AddSummationSign();

        /// <summary>
        /// Add equals sign to the formula.
        /// </summary>
        IFormulaBuilder AddEqualsSign();

        /// <summary>
        /// Add constant sign to the formula.
        /// </summary>
        IFormulaBuilder AddConstantSign();

        /// <summary>
        /// Build the formula
        /// </summary>
        /// <returns>String containing the formula</returns>
        string Build();
    }
}