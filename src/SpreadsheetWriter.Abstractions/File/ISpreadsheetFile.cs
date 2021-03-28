using System;
using System.Threading.Tasks;

namespace SpreadsheetWriter.Abstractions.File
{
    /// <summary>
    /// Abstraction for spreadsheet files.
    /// </summary>
    public interface ISpreadsheetFile : IDisposable
    {
        /// <summary>
        /// Gets the writer from the spreadsheet file.
        /// </summary>
        ISpreadsheetWriter GetSpreadsheetWriter();

        /// <summary>
        /// Saves the spreadsheet.
        /// </summary>
        Task<SaveResult> SaveAsync();
    }
}