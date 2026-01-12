using System;
using System.IO;

using ClosedXML.Excel;

namespace FluentClosedXml
{

    /// <summary>
    /// Fluent API wrapper for ClosedXML workbook operations
    /// </summary>
    public class FluentWorkbook : IDisposable
    {
        private readonly XLWorkbook _workbook;
        private bool _disposed = false;

        /// <summary>
        /// Creates a new fluent workbook wrapper
        /// </summary>
        public FluentWorkbook()
        {
            _workbook = new XLWorkbook();
        }

        /// <summary>
        /// Creates a fluent workbook wrapper from an existing file
        /// </summary>
        /// <param name="filePath">Path to the Excel file</param>
        public FluentWorkbook(string filePath)
        {
            _workbook = new XLWorkbook(filePath);
        }

        /// <summary>
        /// Creates a fluent workbook wrapper from a stream
        /// </summary>
        /// <param name="stream">Stream containing Excel data</param>
        public FluentWorkbook(Stream stream)
        {
            _workbook = new XLWorkbook(stream);
        }

        /// <summary>
        /// Gets the underlying ClosedXML workbook for advanced operations
        /// </summary>
        public XLWorkbook Workbook => _workbook;

        /// <summary>
        /// Creates a new worksheet with fluent API
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <returns>FluentWorksheet for chaining operations</returns>
        public FluentWorksheet AddWorksheet(string name = "Sheet1")
        {
            var worksheet = _workbook.Worksheets.Add(name);
            return new FluentWorksheet(worksheet);
        }

        /// <summary>
        /// Gets an existing worksheet by name with fluent API
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <returns>FluentWorksheet for chaining operations</returns>
        public FluentWorksheet GetWorksheet(string name)
        {
            var worksheet = _workbook.Worksheet(name);
            return new FluentWorksheet(worksheet);
        }

        /// <summary>
        /// Gets an existing worksheet by index with fluent API
        /// </summary>
        /// <param name="index">Index of the worksheet (1-based)</param>
        /// <returns>FluentWorksheet for chaining operations</returns>
        public FluentWorksheet GetWorksheet(int index)
        {
            var worksheet = _workbook.Worksheet(index);
            return new FluentWorksheet(worksheet);
        }

        /// <summary>
        /// Saves the workbook to a file
        /// </summary>
        /// <param name="filePath">Path where to save the file</param>
        /// <returns>This FluentWorkbook for method chaining</returns>
        public FluentWorkbook SaveAs(string filePath)
        {
            _workbook.SaveAs(filePath);
            return this;
        }

        /// <summary>
        /// Saves the workbook to a stream
        /// </summary>
        /// <param name="stream">Stream to save to</param>
        /// <returns>This FluentWorkbook for method chaining</returns>
        public FluentWorkbook SaveAs(Stream stream)
        {
            _workbook.SaveAs(stream);
            return this;
        }

        /// <summary>
        /// Sets workbook properties using a fluent interface
        /// </summary>
        /// <param name="configure">Action to configure workbook properties</param>
        /// <returns>This FluentWorkbook for method chaining</returns>
        public FluentWorkbook WithProperties(Action<IXLWorkbook> configure)
        {
            configure(_workbook);
            return this;
        }

        /// <summary>
        /// Static factory method to create a new fluent workbook
        /// </summary>
        /// <returns>New FluentWorkbook instance</returns>
        public static FluentWorkbook Create() => new FluentWorkbook();

        /// <summary>
        /// Static factory method to load an existing workbook
        /// </summary>
        /// <param name="filePath">Path to the Excel file</param>
        /// <returns>FluentWorkbook instance</returns>
        public static FluentWorkbook Load(string filePath) => new FluentWorkbook(filePath);

        /// <summary>
        /// Static factory method to load a workbook from a stream
        /// </summary>
        /// <param name="stream">Stream containing Excel data</param>
        /// <returns>FluentWorkbook instance</returns>
        public static FluentWorkbook Load(Stream stream) => new FluentWorkbook(stream);

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed && disposing)
            {
                _workbook?.Dispose();
                _disposed = true;
            }
        }
    }
}
