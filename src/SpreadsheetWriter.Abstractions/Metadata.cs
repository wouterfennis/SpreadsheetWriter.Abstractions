using System;

namespace SpreadsheetWriter.Abstractions
{
    /// <summary>
    /// Metadata for the spreadsheet.
    /// </summary>
    public class Metadata
    {
        /// <summary>
        /// The author of the spreadsheet.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// The title of the spreadsheet.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// The subject of the spreadsheet.
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// The date the spreadsheet was created.
        /// </summary>
        public DateTime Created { get; set; }

        /// <summary>
        /// The filename of the spreadsheet.
        /// </summary>
        public string FileName { get; set; }
    }
}