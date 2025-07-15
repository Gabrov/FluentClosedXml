using ClosedXML.Excel;

namespace FluentClosedXml;

/// <summary>
/// Fluent API wrapper for ClosedXML worksheet operations
/// </summary>
public class FluentWorksheet
{
    private readonly IXLWorksheet _worksheet;

    internal FluentWorksheet(IXLWorksheet worksheet)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
    }

    /// <summary>
    /// Gets the underlying ClosedXML worksheet for advanced operations
    /// </summary>
    public IXLWorksheet Worksheet => _worksheet;

    /// <summary>
    /// Sets a value in a cell using fluent API
    /// </summary>
    /// <param name="cellAddress">Cell address (e.g., "A1")</param>
    /// <param name="value">Value to set</param>
    /// <returns>FluentCell for further cell operations</returns>
    public FluentCell SetCell(string cellAddress, object? value)
    {
        var cell = _worksheet.Cell(cellAddress);
        if (value != null)
        {
            cell.Value = XLCellValue.FromObject(value);
        }
        return new FluentCell(cell);
    }

    /// <summary>
    /// Sets a value in a cell using row and column coordinates
    /// </summary>
    /// <param name="row">Row number (1-based)</param>
    /// <param name="column">Column number (1-based)</param>
    /// <param name="value">Value to set</param>
    /// <returns>FluentCell for further cell operations</returns>
    public FluentCell SetCell(int row, int column, object? value)
    {
        var cell = _worksheet.Cell(row, column);
        if (value != null)
        {
            cell.Value = XLCellValue.FromObject(value);
        }
        return new FluentCell(cell);
    }

    /// <summary>
    /// Sets a formula in a cell using fluent API
    /// </summary>
    /// <param name="cellAddress">Cell address (e.g., "A1")</param>
    /// <param name="formula">Formula to set (without = sign)</param>
    /// <returns>FluentCell for further cell operations</returns>
    public FluentCell SetFormula(string cellAddress, string formula)
    {
        var cell = _worksheet.Cell(cellAddress);
        cell.FormulaA1 = formula;
        return new FluentCell(cell);
    }

    /// <summary>
    /// Sets a formula in a cell using row and column coordinates
    /// </summary>
    /// <param name="row">Row number (1-based)</param>
    /// <param name="column">Column number (1-based)</param>
    /// <param name="formula">Formula to set (without = sign)</param>
    /// <returns>FluentCell for further cell operations</returns>
    public FluentCell SetFormula(int row, int column, string formula)
    {
        var cell = _worksheet.Cell(row, column);
        cell.FormulaA1 = formula;
        return new FluentCell(cell);
    }

    /// <summary>
    /// Sets a SUM formula in a cell
    /// </summary>
    /// <param name="cellAddress">Cell address (e.g., "A1")</param>
    /// <param name="range">Range to sum (e.g., "A1:A10")</param>
    /// <returns>FluentCell for further cell operations</returns>
    public FluentCell SetSum(string cellAddress, string range)
    {
        return SetFormula(cellAddress, $"SUM({range})");
    }

    /// <summary>
    /// Sets an AVERAGE formula in a cell
    /// </summary>
    /// <param name="cellAddress">Cell address (e.g., "A1")</param>
    /// <param name="range">Range to average (e.g., "A1:A10")</param>
    /// <returns>FluentCell for further cell operations</returns>
    public FluentCell SetAverage(string cellAddress, string range)
    {
        return SetFormula(cellAddress, $"AVERAGE({range})");
    }

    /// <summary>
    /// Gets a cell for fluent operations
    /// </summary>
    /// <param name="cellAddress">Cell address (e.g., "A1")</param>
    /// <returns>FluentCell for cell operations</returns>
    public FluentCell GetCell(string cellAddress)
    {
        var cell = _worksheet.Cell(cellAddress);
        return new FluentCell(cell);
    }

    /// <summary>
    /// Gets a cell for fluent operations using row and column coordinates
    /// </summary>
    /// <param name="row">Row number (1-based)</param>
    /// <param name="column">Column number (1-based)</param>
    /// <returns>FluentCell for cell operations</returns>
    public FluentCell GetCell(int row, int column)
    {
        var cell = _worksheet.Cell(row, column);
        return new FluentCell(cell);
    }

    /// <summary>
    /// Gets a range for fluent operations
    /// </summary>
    /// <param name="rangeAddress">Range address (e.g., "A1:B10")</param>
    /// <returns>FluentRange for range operations</returns>
    public FluentRange GetRange(string rangeAddress)
    {
        var range = _worksheet.Range(rangeAddress);
        return new FluentRange(range);
    }

    /// <summary>
    /// Gets a range for fluent operations using coordinates
    /// </summary>
    /// <param name="firstRow">First row (1-based)</param>
    /// <param name="firstColumn">First column (1-based)</param>
    /// <param name="lastRow">Last row (1-based)</param>
    /// <param name="lastColumn">Last column (1-based)</param>
    /// <returns>FluentRange for range operations</returns>
    public FluentRange GetRange(int firstRow, int firstColumn, int lastRow, int lastColumn)
    {
        var range = _worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
        return new FluentRange(range);
    }

    /// <summary>
    /// Adds data as a table with fluent API
    /// </summary>
    /// <param name="data">Enumerable data to add</param>
    /// <param name="startCell">Starting cell address (default: A1)</param>
    /// <param name="createTable">Whether to create an Excel table</param>
    /// <returns>FluentRange representing the data range</returns>
    public FluentRange AddData<T>(IEnumerable<T> data, string startCell = "A1", bool createTable = false)
    {
        var cell = _worksheet.Cell(startCell);
        var range = cell.InsertData(data, transpose: false);
        
        if (createTable)
        {
            range.CreateTable();
        }
        
        return new FluentRange(range);
    }

    /// <summary>
    /// Sets column headers starting from a specific cell
    /// </summary>
    /// <param name="headers">Array of header names</param>
    /// <param name="startCell">Starting cell address (default: A1)</param>
    /// <returns>FluentRange representing the header range</returns>
    public FluentRange AddHeaders(string[] headers, string startCell = "A1")
    {
        var cell = _worksheet.Cell(startCell);
        var endColumn = cell.Address.ColumnNumber + headers.Length - 1;
        
        for (int i = 0; i < headers.Length; i++)
        {
            _worksheet.Cell(cell.Address.RowNumber, cell.Address.ColumnNumber + i).Value = headers[i];
        }
        
        var range = _worksheet.Range(cell.Address.RowNumber, cell.Address.ColumnNumber, 
                                   cell.Address.RowNumber, endColumn);
        return new FluentRange(range);
    }

    /// <summary>
    /// Auto-fits all columns in the worksheet
    /// </summary>
    /// <returns>This FluentWorksheet for method chaining</returns>
    public FluentWorksheet AutoFitColumns()
    {
        _worksheet.Columns().AdjustToContents();
        return this;
    }

    /// <summary>
    /// Auto-fits all rows in the worksheet
    /// </summary>
    /// <returns>This FluentWorksheet for method chaining</returns>
    public FluentWorksheet AutoFitRows()
    {
        _worksheet.Rows().AdjustToContents();
        return this;
    }

    /// <summary>
    /// Sets the worksheet name
    /// </summary>
    /// <param name="name">New worksheet name</param>
    /// <returns>This FluentWorksheet for method chaining</returns>
    public FluentWorksheet WithName(string name)
    {
        _worksheet.Name = name;
        return this;
    }

    /// <summary>
    /// Configures worksheet properties using a fluent interface
    /// </summary>
    /// <param name="configure">Action to configure worksheet properties</param>
    /// <returns>This FluentWorksheet for method chaining</returns>
    public FluentWorksheet WithProperties(Action<IXLWorksheet> configure)
    {
        configure(_worksheet);
        return this;
    }

    /// <summary>
    /// Freezes panes at the specified cell
    /// </summary>
    /// <param name="cellAddress">Cell address where to freeze panes</param>
    /// <returns>This FluentWorksheet for method chaining</returns>
    public FluentWorksheet FreezePanes(string cellAddress)
    {
        _worksheet.SheetView.FreezeRows(_worksheet.Cell(cellAddress).Address.RowNumber - 1);
        _worksheet.SheetView.FreezeColumns(_worksheet.Cell(cellAddress).Address.ColumnNumber - 1);
        return this;
    }

    /// <summary>
    /// Freezes the first row (typically headers)
    /// </summary>
    /// <returns>This FluentWorksheet for method chaining</returns>
    public FluentWorksheet FreezeTopRow()
    {
        _worksheet.SheetView.FreezeRows(1);
        return this;
    }
}