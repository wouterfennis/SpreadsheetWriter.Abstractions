using System.Drawing;
using SpreadsheetWriter.Abstractions.Cell;
using SpreadsheetWriter.Abstractions.Formula;
using SpreadsheetWriter.Abstractions.Styling;

namespace SpreadsheetWriter.Abstractions
{
    /// <inheritdoc/>
    public abstract class SpreadsheetWriterBase : ISpreadsheetWriter
    {
        private const int DefaultTextRotation = 0;
        private const float DefaultFontSize = 11;
        private const bool DefaultIsFontBold = false;
        private const string DefaultFormat = "";
        private const BorderStyle DefaultBorderStyle = BorderStyle.None;
        private const BorderDirection DefaultBorderDirection = BorderDirection.None;

        private readonly Color DefaultBackgroundColor = Color.White;
        private readonly Color DefaultFontColor = Color.Black;
        private readonly Color DefaultBorderColor = Color.LightGray;
        private readonly int DefaultXPosition;
        private readonly int DefaultYPosition;

        protected Color CurrentBackgroundColor;
        protected Color CurrentFontColor;
        protected Color CurrentBorderColor;
        protected float CurrentFontSize;
        protected bool IsCurrentFontBold;
        protected int CurrentTextRotation;
        protected string CurrentFormat;
        protected BorderStyle CurrentBorderStyle;
        protected BorderDirection CurrentBorderDirection;

        /// <summary>
        /// The position of the current selected Cell
        /// </summary>
        public Point CurrentPosition { get; set; }

        public SpreadsheetWriterBase(int defaultXPosition, int defaultYPosition)
        {
            DefaultXPosition = defaultXPosition;
            DefaultYPosition = defaultYPosition;
            CurrentPosition = new Point(DefaultXPosition, DefaultYPosition);
        }

        /// <inheritdoc/>
        public abstract ICellRange GetCellRange(Point position);

        /// <inheritdoc/>
        public ISpreadsheetWriter MoveDown()
        {
            CurrentPosition = new Point(CurrentPosition.X, CurrentPosition.Y + 1);
            return this;
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter MoveDownTimes(int times)
        {
            for (int i = 0; i < times; i++)
            {
                MoveDown();
            }
            return this;
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter MoveUp()
        {
            CurrentPosition = new Point(CurrentPosition.X, CurrentPosition.Y - 1);
            return this;
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter MoveUpTimes(int times)
        {
            for (int i = 0; i < times; i++)
            {
                MoveUp();
            }
            return this;
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter MoveLeft()
        {
            CurrentPosition = new Point(CurrentPosition.X - 1, CurrentPosition.Y);
            return this;
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter MoveLeftTimes(int times)
        {
            for (int i = 0; i < times; i++)
            {
                MoveLeft();
            }
            return this;
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter MoveRight()
        {
            CurrentPosition = new Point(CurrentPosition.X + 1, CurrentPosition.Y);
            return this;
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter MoveRightTimes(int times)
        {
            for (int i = 0; i < times; i++)
            {
                MoveRight();
            }
            return this;
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter SetBackgroundColor(Color color)
        {
            CurrentBackgroundColor = color;
            return this;
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter SetFontColor(Color color)
        {
            CurrentFontColor = color;
            return this;
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter SetTextRotation(int rotation)
        {
            CurrentTextRotation = rotation;
            return this;
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter SetFontSize(float size)
        {
            CurrentFontSize = size;
            return this;
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter SetFontBold(bool isActive)
        {
            IsCurrentFontBold = isActive;
            return this;
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter SetFormat(string format)
        {
            CurrentFormat = format;
            return this;
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter SetBorder(BorderStyle borderStyle, BorderDirection borderDirection, Color borderColor)
        {
            CurrentBorderStyle = borderStyle;
            CurrentBorderDirection = borderDirection;
            CurrentBorderColor = borderColor;
            return this;
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter NewLine()
        {
            CurrentPosition = new Point(DefaultXPosition, CurrentPosition.Y + 1);
            return this;
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter ResetStyling()
        {
            CurrentBackgroundColor = DefaultBackgroundColor;
            CurrentFontColor = DefaultFontColor;
            CurrentBorderColor = DefaultBorderColor;
            CurrentTextRotation = DefaultTextRotation;
            CurrentFontSize = DefaultFontSize;
            IsCurrentFontBold = DefaultIsFontBold;
            CurrentFormat = DefaultFormat;
            CurrentBorderStyle = DefaultBorderStyle;
            CurrentBorderDirection = DefaultBorderDirection;

            return this;
        }

        /// <inheritdoc/>
        public abstract ISpreadsheetWriter Write(decimal value);

        /// <inheritdoc/>
        public abstract ISpreadsheetWriter Write(string value);

        /// <inheritdoc/>
        public abstract ISpreadsheetWriter PlaceStandardFormula(Point startPosition, Point endPosition, FormulaType formulaType);

        /// <inheritdoc/>
        public abstract ISpreadsheetWriter PlaceCustomFormula(IFormulaBuilder formulaBuilder);

    }
}
