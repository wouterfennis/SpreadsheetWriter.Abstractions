using System.Drawing;
using SpreadsheetWriter.Abstractions.Cell;
using SpreadsheetWriter.Abstractions.Formula;
using SpreadsheetWriter.Abstractions.Styling;

namespace SpreadsheetWriter.Abstractions
{
    /// <summary>
    /// Writer to create spreadsheets fluently.
    /// </summary>
    public interface ISpreadsheetWriter
    {
        /// <summary>
        /// The current location of the selected cell in the spreadsheet.
        /// </summary>
        Point CurrentPosition { get; set; }

        /// <summary>
        /// Get the cell range of an point in the spreadsheet.
        /// </summary>
        ICellRange GetCellRange(Point position);

        /// <summary>
        /// Move one position up from the current selected cell.
        /// </summary>
        ISpreadsheetWriter MoveUp();

        /// <summary>
        /// Move a number of positions up from the current selected cell.
        /// </summary>
        ISpreadsheetWriter MoveUpTimes(int times);

        /// <summary>
        /// Move one position down from the current selected cell.
        /// </summary>
        ISpreadsheetWriter MoveDown();

        /// <summary>
        /// Move a number of positions down from the current selected cell.
        /// </summary>
        ISpreadsheetWriter MoveDownTimes(int times);

        /// <summary>
        /// Move one position to the left from the current selected cell.
        /// </summary>
        ISpreadsheetWriter MoveLeft();

        /// <summary>
        /// Move a number positions to the left from the current selected cell.
        /// </summary>
        ISpreadsheetWriter MoveLeftTimes(int times);

        /// <summary>
        /// Move one position to the right from the current selected cell.
        /// </summary>
        ISpreadsheetWriter MoveRight();

        /// <summary>
        /// Move a number positions to the right from the current selected cell.
        /// </summary>
        ISpreadsheetWriter MoveRightTimes(int times);

        /// <summary>
        /// Set the background color for future inserts to a specified <see cref="Color"/>.
        /// </summary>
        ISpreadsheetWriter SetBackgroundColor(Color color);

        /// <summary>
        /// Set the font color for future inserts to a specified <see cref="Color"/>.
        /// </summary>
        ISpreadsheetWriter SetFontColor(Color color);

        /// <summary>
        /// Set the font size for future inserts to a specified size.
        /// </summary>
        ISpreadsheetWriter SetFontSize(float size);

        /// <summary>
        /// Set the font bold for future inserts to a specified size.
        /// </summary>
        ISpreadsheetWriter SetFontBold(bool isActive);

        /// <summary>
        /// Set the text rotation for future inserts to a specified rotation.
        /// </summary>
        ISpreadsheetWriter SetTextRotation(int rotation);

        /// <summary>
        /// Set the format for future inserts to a specified format.
        /// </summary>
        ISpreadsheetWriter SetFormat(string format);

        /// <summary>
        /// Set the border for future inserts to a specified style and direction.
        /// </summary>
        ISpreadsheetWriter SetBorder(BorderStyle borderStyle, BorderDirection borderDirection, Color borderColor);

        /// <summary>
        /// Move the current selected cell to the start of one row below.
        /// </summary>
        ISpreadsheetWriter NewLine();

        /// <summary>
        /// Resets any configured styling to the default values.
        /// </summary>
        ISpreadsheetWriter ResetStyling();

        /// <summary>
        /// Write a decimal value in the current selected cell.
        /// </summary>
        ISpreadsheetWriter Write(decimal value);

        /// <summary>
        /// Write a string value in the current selected cell.
        /// </summary>
        ISpreadsheetWriter Write(string value);

        /// <summary>
        /// Place the outcome of a standard formula in the spreadsheet at the current position.
        /// </summary>
        /// <param name="startPosition">The position where input for the formula should start.</param>
        /// <param name="endPosition">The position where the input for the formula should end.</param>
        /// <param name="formulaType">The type of formula to consume the data.</param>
        ISpreadsheetWriter PlaceStandardFormula(Point startPosition, Point endPosition, FormulaType formulaType);

        /// <summary>
        /// Place the outcome of a custom formula in the spreadsheet at the current position.
        /// </summary>
        /// <param name="customFormulaBuilder">The configured formula builder.</param>
        ISpreadsheetWriter PlaceCustomFormula(IFormulaBuilder customFormulaBuilder);
    }
}
