using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FluentClosedXml;

/// <summary>
/// Fluent API wrapper for ClosedXML cell operations
/// </summary>
public class FluentCell
{
    private readonly IXLCell _cell;

    internal FluentCell(IXLCell cell)
    {
        _cell = cell ?? throw new ArgumentNullException(nameof(cell));
    }

    /// <summary>
    /// Gets the underlying ClosedXML cell for advanced operations
    /// </summary>
    public IXLCell Cell => _cell;

    /// <summary>
    /// Gets the range address as a string (e.g., "A1:B10")
    /// </summary>
    /// <returns>The range address in A1 notation</returns>
    public string GetAddress()
    {
        return _cell.Address.ToString();
    }

    /// <summary>
    /// Sets the cell value
    /// </summary>
    /// <param name="value">Value to set</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithValue(object? value)
    {
        if (value != null)
        {
            _cell.Value = XLCellValue.FromObject(value);
        }
        return this;
    }

    /// <summary>
    /// Sets the cell formula using A1 notation
    /// </summary>
    /// <param name="formula">Formula to set (without = sign)</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithFormula(string formula)
    {
        _cell.FormulaA1 = formula;
        return this;
    }

    /// <summary>
    /// Sets the cell formula using R1C1 notation
    /// </summary>
    /// <param name="formula">Formula in R1C1 notation (without = sign)</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithFormulaR1C1(string formula)
    {
        _cell.FormulaR1C1 = formula;
        return this;
    }

    #region Common Formula Methods

    /// <summary>
    /// Sets a SUM formula for the specified range
    /// </summary>
    /// <param name="range">Range to sum (e.g., "A1:A10")</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithSum(string range)
    {
        _cell.FormulaA1 = $"SUM({range})";
        return this;
    }

    /// <summary>
    /// Sets a SUM formula for multiple ranges with optional negative signs
    /// </summary>
    /// <param name="ranges">Multiple ranges to sum, with optional negative signs (e.g., "A1:A5", "-B21", "C1:C10")</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithSum(params string[] ranges)
    {
        if (ranges == null || ranges.Length == 0)
            throw new ArgumentException("At least one range must be provided", nameof(ranges));

        if (ranges.Length == 1)
        {
            // Use the single-range version for efficiency
            return WithSum(ranges[0]);
        }

        var formulaParts = new List<string>();
        
        foreach (var range in ranges)
        {
            if (string.IsNullOrWhiteSpace(range))
                continue;

            var trimmedRange = range.Trim();
            
            if (trimmedRange.StartsWith("-"))
            {
                // Handle negative range: -B21 becomes -SUM(B21) or -B21 if it's a single cell
                var rangeWithoutMinus = trimmedRange.Substring(1).Trim();
                
                // Check if it's a single cell reference or a range
                if (IsSingleCellReference(rangeWithoutMinus))
                {
                    formulaParts.Add($"-{rangeWithoutMinus}");
                }
                else
                {
                    formulaParts.Add($"-SUM({rangeWithoutMinus})");
                }
            }
            else
            {
                // Handle positive range: A1:A5 becomes SUM(A1:A5) or A1 if it's a single cell
                if (IsSingleCellReference(trimmedRange))
                {
                    formulaParts.Add(trimmedRange);
                }
                else
                {
                    formulaParts.Add($"SUM({trimmedRange})");
                }
            }
        }

        if (formulaParts.Count == 0)
            throw new ArgumentException("No valid ranges provided", nameof(ranges));

        // Create the final formula
        var formula = string.Join("+", formulaParts);
        _cell.FormulaA1 = formula;
        return this;
    }

    /// <summary>
    /// Helper method to determine if a range reference is a single cell
    /// </summary>
    /// <param name="range">Range reference to check</param>
    /// <returns>True if it's a single cell reference, false if it's a range</returns>
    private static bool IsSingleCellReference(string range)
    {
        if (string.IsNullOrWhiteSpace(range))
            return false;

        // Single cell references don't contain ':' 
        // Examples: A1, B21, $A$1, Sheet1!A1 are single cells
        // Examples: A1:A5, B1:C10 are ranges
        return !range.Contains(':');
    }

    /// <summary>
    /// Sets an AVERAGE formula for the specified range
    /// </summary>
    /// <param name="range">Range to average (e.g., "A1:A10")</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithAverage(string range)
    {
        _cell.FormulaA1 = $"AVERAGE({range})";
        return this;
    }

    /// <summary>
    /// Sets a COUNT formula for the specified range
    /// </summary>
    /// <param name="range">Range to count (e.g., "A1:A10")</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithCount(string range)
    {
        _cell.FormulaA1 = $"COUNT({range})";
        return this;
    }

    /// <summary>
    /// Sets a COUNTA formula for the specified range (counts non-empty cells)
    /// </summary>
    /// <param name="range">Range to count (e.g., "A1:A10")</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithCountA(string range)
    {
        _cell.FormulaA1 = $"COUNTA({range})";
        return this;
    }

    /// <summary>
    /// Sets a MAX formula for the specified range
    /// </summary>
    /// <param name="range">Range to find maximum (e.g., "A1:A10")</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithMax(string range)
    {
        _cell.FormulaA1 = $"MAX({range})";
        return this;
    }

    /// <summary>
    /// Sets a MIN formula for the specified range
    /// </summary>
    /// <param name="range">Range to find minimum (e.g., "A1:A10")</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithMin(string range)
    {
        _cell.FormulaA1 = $"MIN({range})";
        return this;
    }

    /// <summary>
    /// Sets a VLOOKUP formula
    /// </summary>
    /// <param name="lookupValue">Value to lookup</param>
    /// <param name="tableArray">Table array range</param>
    /// <param name="columnIndex">Column index to return</param>
    /// <param name="exactMatch">Whether to use exact match (default: true)</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithVLookup(string lookupValue, string tableArray, int columnIndex, bool exactMatch = true)
    {
        var matchType = exactMatch ? "FALSE" : "TRUE";
        _cell.FormulaA1 = $"VLOOKUP({lookupValue},{tableArray},{columnIndex},{matchType})";
        return this;
    }

    /// <summary>
    /// Sets an IF formula
    /// </summary>
    /// <param name="condition">Condition to test</param>
    /// <param name="valueIfTrue">Value if condition is true</param>
    /// <param name="valueIfFalse">Value if condition is false</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithIf(string condition, string valueIfTrue, string valueIfFalse)
    {
        _cell.FormulaA1 = $"IF({condition},{valueIfTrue},{valueIfFalse})";
        return this;
    }

    /// <summary>
    /// Sets a CONCATENATE formula
    /// </summary>
    /// <param name="values">Values to concatenate</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithConcatenate(params string[] values)
    {
        var concatenatedValues = string.Join(",", values);
        _cell.FormulaA1 = $"CONCATENATE({concatenatedValues})";
        return this;
    }

    /// <summary>
    /// Sets a SUMIF formula
    /// </summary>
    /// <param name="range">Range to evaluate</param>
    /// <param name="criteria">Criteria for sum</param>
    /// <param name="sumRange">Range to sum (optional, uses range if not specified)</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithSumIf(string range, string criteria, string? sumRange = null)
    {
        if (string.IsNullOrEmpty(sumRange))
        {
            _cell.FormulaA1 = $"SUMIF({range},{criteria})";
        }
        else
        {
            _cell.FormulaA1 = $"SUMIF({range},{criteria},{sumRange})";
        }
        return this;
    }

    /// <summary>
    /// Sets a COUNTIF formula
    /// </summary>
    /// <param name="range">Range to evaluate</param>
    /// <param name="criteria">Criteria for counting</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithCountIf(string range, string criteria)
    {
        _cell.FormulaA1 = $"COUNTIF({range},{criteria})";
        return this;
    }

    /// <summary>
    /// Sets a TODAY formula
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithToday()
    {
        _cell.FormulaA1 = "TODAY()";
        return this;
    }

    /// <summary>
    /// Sets a NOW formula
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithNow()
    {
        _cell.FormulaA1 = "NOW()";
        return this;
    }

    /// <summary>
    /// Sets a ROUND formula
    /// </summary>
    /// <param name="value">Value to round</param>
    /// <param name="decimals">Number of decimal places</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithRound(string value, int decimals)
    {
        _cell.FormulaA1 = $"ROUND({value},{decimals})";
        return this;
    }

    /// <summary>
    /// Sets a simple arithmetic formula between two cells
    /// </summary>
    /// <param name="cell1">First cell reference</param>
    /// <param name="operation">Operation (+, -, *, /)</param>
    /// <param name="cell2">Second cell reference</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithArithmetic(string cell1, string operation, string cell2)
    {
        _cell.FormulaA1 = $"{cell1}{operation}{cell2}";
        return this;
    }

    /// <summary>
    /// Sets a formula that adds two cell references
    /// </summary>
    /// <param name="cell1">First cell reference</param>
    /// <param name="cell2">Second cell reference</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithAdd(string cell1, string cell2)
    {
        return WithArithmetic(cell1, "+", cell2);
    }

    /// <summary>
    /// Sets a formula that subtracts two cell references
    /// </summary>
    /// <param name="cell1">First cell reference</param>
    /// <param name="cell2">Second cell reference</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithSubtract(string cell1, string cell2)
    {
        return WithArithmetic(cell1, "-", cell2);
    }

    /// <summary>
    /// Sets a formula that multiplies two cell references
    /// </summary>
    /// <param name="cell1">First cell reference</param>
    /// <param name="cell2">Second cell reference</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithMultiply(string cell1, string cell2)
    {
        return WithArithmetic(cell1, "*", cell2);
    }

    /// <summary>
    /// Sets a formula that divides two cell references
    /// </summary>
    /// <param name="cell1">First cell reference</param>
    /// <param name="cell2">Second cell reference</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithDivide(string cell1, string cell2)
    {
        return WithArithmetic(cell1, "/", cell2);
    }

    #endregion

    #region Formula Utilities

    /// <summary>
    /// Gets the formula from the cell
    /// </summary>
    /// <returns>The cell's formula in A1 notation</returns>
    public string GetFormula()
    {
        return _cell.FormulaA1;
    }

    /// <summary>
    /// Gets the formula from the cell in R1C1 notation
    /// </summary>
    /// <returns>The cell's formula in R1C1 notation</returns>
    public string GetFormulaR1C1()
    {
        return _cell.FormulaR1C1;
    }

    /// <summary>
    /// Checks if the cell contains a formula
    /// </summary>
    /// <returns>True if the cell has a formula</returns>
    public bool HasFormula()
    {
        return _cell.HasFormula;
    }

    /// <summary>
    /// Gets the calculated value of the formula
    /// </summary>
    /// <returns>The calculated result of the formula</returns>
    public object? GetFormulaResult()
    {
        return _cell.CachedValue;
    }

    #endregion

    #region Predefined Number Format Methods

    /// <summary>
    /// Sets the number format using ClosedXML's predefined format
    /// </summary>
    /// <param name="formatId">Predefined format ID</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithPredefinedFormat(int formatId)
    {
        _cell.Style.NumberFormat.NumberFormatId = formatId;
        return this;
    }

    /// <summary>
    /// Applies General format (Excel's default)
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsGeneral()
    {
        _cell.Style.NumberFormat.NumberFormatId = 0;
        return this;
    }

    /// <summary>
    /// Applies format: 0 (Number with no decimals)
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsNumber()
    {
        _cell.Style.NumberFormat.NumberFormatId = 1;
        return this;
    }

    /// <summary>
    /// Applies format: 0.00 (Number with 2 decimals)
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsNumberWith2Decimals()
    {
        _cell.Style.NumberFormat.NumberFormatId = 2;
        return this;
    }

    /// <summary>
    /// Applies format: #,##0 (Number with thousands separator)
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsNumberWithThousandsSeparator()
    {
        _cell.Style.NumberFormat.NumberFormatId = 3;
        return this;
    }

    /// <summary>
    /// Applies format: #,##0.00 (Number with thousands separator and 2 decimals)
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsNumberWithThousandsSeparatorAnd2Decimals()
    {
        _cell.Style.NumberFormat.NumberFormatId = 4;
        return this;
    }

    /// <summary>
    /// Applies standard currency format
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsCurrencyBuiltIn()
    {
        _cell.Style.NumberFormat.NumberFormatId = 5;
        return this;
    }

    /// <summary>
    /// Applies currency format with red negative values
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsCurrencyWithRedNegatives()
    {
        _cell.Style.NumberFormat.NumberFormatId = 6;
        return this;
    }

    /// <summary>
    /// Applies currency format with 2 decimals
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsCurrencyWith2Decimals()
    {
        _cell.Style.NumberFormat.NumberFormatId = 7;
        return this;
    }

    /// <summary>
    /// Applies currency format with 2 decimals and red negative values
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsCurrencyWith2DecimalsAndRedNegatives()
    {
        _cell.Style.NumberFormat.NumberFormatId = 8;
        return this;
    }

    /// <summary>
    /// Applies percentage format: 0%
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsPercentageBuiltIn()
    {
        _cell.Style.NumberFormat.NumberFormatId = 9;
        return this;
    }

    /// <summary>
    /// Applies percentage format with 2 decimals: 0.00%
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsPercentageWith2Decimals()
    {
        _cell.Style.NumberFormat.NumberFormatId = 10;
        return this;
    }

    /// <summary>
    /// Applies scientific notation format: 0.00E+00
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsScientific()
    {
        _cell.Style.NumberFormat.NumberFormatId = 11;
        return this;
    }

    /// <summary>
    /// Applies fraction format: # ?/?
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsFraction()
    {
        _cell.Style.NumberFormat.NumberFormatId = 12;
        return this;
    }

    /// <summary>
    /// Applies fraction format with denominators up to 99: # ??/??
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsFractionUpTo99()
    {
        _cell.Style.NumberFormat.NumberFormatId = 13;
        return this;
    }

    /// <summary>
    /// Applies short date format: mm/dd/yyyy
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsShortDate()
    {
        _cell.Style.NumberFormat.NumberFormatId = 14;
        return this;
    }

    /// <summary>
    /// Applies date format: d-mmm-yy
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsDateDMmmYy()
    {
        _cell.Style.NumberFormat.NumberFormatId = 15;
        return this;
    }

    /// <summary>
    /// Applies date format: d-mmm
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsDateDMmm()
    {
        _cell.Style.NumberFormat.NumberFormatId = 16;
        return this;
    }

    /// <summary>
    /// Applies date format: mmm-yy
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsDateMmmYy()
    {
        _cell.Style.NumberFormat.NumberFormatId = 17;
        return this;
    }

    /// <summary>
    /// Applies time format: h:mm AM/PM
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsTimeAmPm()
    {
        _cell.Style.NumberFormat.NumberFormatId = 18;
        return this;
    }

    /// <summary>
    /// Applies time format with seconds: h:mm:ss AM/PM
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsTimeWithSecondsAmPm()
    {
        _cell.Style.NumberFormat.NumberFormatId = 19;
        return this;
    }

    /// <summary>
    /// Applies 24-hour time format: h:mm
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsTime24Hour()
    {
        _cell.Style.NumberFormat.NumberFormatId = 20;
        return this;
    }

    /// <summary>
    /// Applies 24-hour time format with seconds: h:mm:ss
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsTime24HourWithSeconds()
    {
        _cell.Style.NumberFormat.NumberFormatId = 21;
        return this;
    }

    /// <summary>
    /// Applies date and time format: mm/dd/yyyy h:mm
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsDateTime()
    {
        _cell.Style.NumberFormat.NumberFormatId = 22;
        return this;
    }

    /// <summary>
    /// Applies accounting format with no decimals
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsAccounting()
    {
        _cell.Style.NumberFormat.NumberFormatId = 37;
        return this;
    }

    /// <summary>
    /// Applies accounting format with 2 decimals
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsAccountingWith2Decimals()
    {
        _cell.Style.NumberFormat.NumberFormatId = 38;
        return this;
    }

    /// <summary>
    /// Applies accounting format with red negatives
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsAccountingWithRedNegatives()
    {
        _cell.Style.NumberFormat.NumberFormatId = 39;
        return this;
    }

    /// <summary>
    /// Applies accounting format with 2 decimals and red negatives
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsAccountingWith2DecimalsAndRedNegatives()
    {
        _cell.Style.NumberFormat.NumberFormatId = 40;
        return this;
    }

    /// <summary>
    /// Applies text format (displays numbers as text)
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsText()
    {
        _cell.Style.NumberFormat.NumberFormatId = 49;
        return this;
    }

    /// <summary>
    /// Sets format for elapsed time in hours: [h]:mm:ss
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsElapsedTimeHours()
    {
        _cell.Style.NumberFormat.Format = "[h]:mm:ss";
        return this;
    }

    /// <summary>
    /// Sets format for elapsed time in minutes: [mm]:ss
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsElapsedTimeMinutes()
    {
        _cell.Style.NumberFormat.Format = "[mm]:ss";
        return this;
    }

    /// <summary>
    /// Sets format for elapsed time in seconds: [ss]
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsElapsedTimeSeconds()
    {
        _cell.Style.NumberFormat.Format = "[ss]";
        return this;
    }

    /// <summary>
    /// Applies number format with specific decimal places
    /// </summary>
    /// <param name="decimalPlaces">Number of decimal places (0-30)</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsNumberWithDecimals(int decimalPlaces)
    {
        if (decimalPlaces < 0 || decimalPlaces > 30)
            throw new ArgumentOutOfRangeException(nameof(decimalPlaces), "Decimal places must be between 0 and 30");

        var format = decimalPlaces == 0 ? "0" : "0." + new string('0', decimalPlaces);
        _cell.Style.NumberFormat.Format = format;
        return this;
    }

    /// <summary>
    /// Applies number format with thousands separator and specific decimal places
    /// </summary>
    /// <param name="decimalPlaces">Number of decimal places (0-30)</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsNumberWithThousandsSeparatorAndDecimals(int decimalPlaces)
    {
        if (decimalPlaces < 0 || decimalPlaces > 30)
            throw new ArgumentOutOfRangeException(nameof(decimalPlaces), "Decimal places must be between 0 and 30");

        var format = decimalPlaces == 0 ? "#,##0" : "#,##0." + new string('0', decimalPlaces);
        _cell.Style.NumberFormat.Format = format;
        return this;
    }

    /// <summary>
    /// Applies percentage format with specific decimal places
    /// </summary>
    /// <param name="decimalPlaces">Number of decimal places (0-30)</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsPercentageWithDecimals(int decimalPlaces)
    {
        if (decimalPlaces < 0 || decimalPlaces > 30)
            throw new ArgumentOutOfRangeException(nameof(decimalPlaces), "Decimal places must be between 0 and 30");

        var format = decimalPlaces == 0 ? "0%" : "0." + new string('0', decimalPlaces) + "%";
        _cell.Style.NumberFormat.Format = format;
        return this;
    }

    /// <summary>
    /// Applies currency format with specific decimal places and custom symbol
    /// </summary>
    /// <param name="decimalPlaces">Number of decimal places (0-30)</param>
    /// <param name="currencySymbol">Currency symbol (default: $)</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsCurrencyWithDecimals(int decimalPlaces, string currencySymbol = "$")
    {
        if (decimalPlaces < 0 || decimalPlaces > 30)
            throw new ArgumentOutOfRangeException(nameof(decimalPlaces), "Decimal places must be between 0 and 30");

        var decimalsFormat = decimalPlaces == 0 ? "" : "." + new string('0', decimalPlaces);
        var format = $"\"{currencySymbol}\"#,##0{decimalsFormat}_);(\"{currencySymbol}\"#,##0{decimalsFormat})";
        _cell.Style.NumberFormat.Format = format;
        return this;
    }

    /// <summary>
    /// Gets the current number format ID of the cell
    /// </summary>
    /// <returns>The number format ID</returns>
    public int GetNumberFormatId()
    {
        return _cell.Style.NumberFormat.NumberFormatId;
    }

    /// <summary>
    /// Gets the current number format string of the cell
    /// </summary>
    /// <returns>The number format string</returns>
    public string GetNumberFormat()
    {
        return _cell.Style.NumberFormat.Format;
    }

    #endregion

    /// <summary>
    /// Sets the cell style using a fluent interface
    /// </summary>
    /// <param name="styleAction">Action to configure cell style</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithStyle(Action<IXLStyle> styleAction)
    {
        styleAction(_cell.Style);
        return this;
    }

    /// <summary>
    /// Makes the cell bold
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell Bold()
    {
        _cell.Style.Font.Bold = true;
        return this;
    }

    /// <summary>
    /// Makes the cell italic
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell Italic()
    {
        _cell.Style.Font.Italic = true;
        return this;
    }

    /// <summary>
    /// Sets the font size
    /// </summary>
    /// <param name="size">Font size</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithFontSize(double size)
    {
        _cell.Style.Font.FontSize = size;
        return this;
    }

    /// <summary>
    /// Sets the font color
    /// </summary>
    /// <param name="color">Font color</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithFontColor(XLColor color)
    {
        _cell.Style.Font.FontColor = color;
        return this;
    }

    /// <summary>
    /// Sets the background color
    /// </summary>
    /// <param name="color">Background color</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithBackgroundColor(XLColor color)
    {
        _cell.Style.Fill.BackgroundColor = color;
        return this;
    }

    /// <summary>
    /// Sets horizontal alignment
    /// </summary>
    /// <param name="alignment">Horizontal alignment</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithHorizontalAlignment(XLAlignmentHorizontalValues alignment)
    {
        _cell.Style.Alignment.Horizontal = alignment;
        return this;
    }

    /// <summary>
    /// Sets vertical alignment
    /// </summary>
    /// <param name="alignment">Vertical alignment</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithVerticalAlignment(XLAlignmentVerticalValues alignment)
    {
        _cell.Style.Alignment.Vertical = alignment;
        return this;
    }

    /// <summary>
    /// Centers the cell content both horizontally and vertically
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell Center()
    {
        _cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        _cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        return this;
    }

    /// <summary>
    /// Enables text wrapping for the cell
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithTextWrap()
    {
        _cell.Style.Alignment.WrapText = true;
        return this;
    }

    /// <summary>
    /// Disables text wrapping for the cell
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithoutTextWrap()
    {
        _cell.Style.Alignment.WrapText = false;
        return this;
    }

    /// <summary>
    /// Sets text wrapping for the cell
    /// </summary>
    /// <param name="wrapText">True to enable text wrapping, false to disable</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithTextWrap(bool wrapText)
    {
        _cell.Style.Alignment.WrapText = wrapText;
        return this;
    }

    /// <summary>
    /// Sets borders around the cell
    /// </summary>
    /// <param name="borderStyle">Border style</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithBorder(XLBorderStyleValues borderStyle = XLBorderStyleValues.Thin)
    {
        _cell.Style.Border.OutsideBorder = borderStyle;
        return this;
    }

    /// <summary>
    /// Sets a specific border
    /// </summary>
    /// <param name="borderType">Type of border</param>
    /// <param name="borderStyle">Border style</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithBorder(XLBorderStyleValues borderStyle, params string[] sides)
    {
        foreach (var side in sides)
        {
            switch (side.ToLower())
            {
                case "top":
                    _cell.Style.Border.TopBorder = borderStyle;
                    break;
                case "bottom":
                    _cell.Style.Border.BottomBorder = borderStyle;
                    break;
                case "left":
                    _cell.Style.Border.LeftBorder = borderStyle;
                    break;
                case "right":
                    _cell.Style.Border.RightBorder = borderStyle;
                    break;
            }
        }
        return this;
    }

    /// <summary>
    /// Sets the number format
    /// </summary>
    /// <param name="format">Number format string</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithNumberFormat(string format)
    {
        _cell.Style.NumberFormat.Format = format;
        return this;
    }

    /// <summary>
    /// Formats the cell as currency
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsCurrency()
    {
        _cell.Style.NumberFormat.Format = "\"$\"#,##0.00_);(\"$\"#,##0.00)";
        return this;
    }

    /// <summary>
    /// Formats the cell as percentage
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsPercentage()
    {
        _cell.Style.NumberFormat.Format = "0.00%";
        return this;
    }

    /// <summary>
    /// Formats the cell as date
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell AsDate()
    {
        _cell.Style.NumberFormat.Format = "mm/dd/yyyy";
        return this;
    }

    /// <summary>
    /// Gets the cell value as a specific type
    /// </summary>
    /// <typeparam name="T">Type to convert to</typeparam>
    /// <returns>Cell value as specified type</returns>
    public T GetValue<T>()
    {
        return _cell.GetValue<T>();
    }

    /// <summary>
    /// Gets the cell value as string
    /// </summary>
    /// <returns>Cell value as string</returns>
    public string GetText()
    {
        return _cell.GetText();
    }

    #region Conditional Formatting Methods

    /// <summary>
    /// Applies conditional formatting to the cell
    /// </summary>
    /// <param name="configure">Action to configure conditional formatting</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell WithConditionalFormatting(Action<IXLConditionalFormat> configure)
    {
        var range = _cell.AsRange();
        var conditionalFormat = range.AddConditionalFormat();
        configure(conditionalFormat);
        return this;
    }

    /// <summary>
    /// Highlights the cell if it contains specific text
    /// </summary>
    /// <param name="text">Text to search for</param>
    /// <param name="backgroundColor">Background color for matching cell</param>
    /// <param name="fontColor">Font color for matching cell (optional)</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell HighlightIfContains(string text, XLColor backgroundColor, XLColor? fontColor = null)
    {
        var range = _cell.AsRange();
        var conditionalFormat = range.AddConditionalFormat();
        var format = conditionalFormat.WhenContains(text);
        format.Fill.SetBackgroundColor(backgroundColor);
        
        if (fontColor != null)
        {
            format.Font.SetFontColor(fontColor);
        }
        
        return this;
    }

    /// <summary>
    /// Highlights the cell if it is greater than a specific value
    /// </summary>
    /// <param name="value">Value to compare against</param>
    /// <param name="backgroundColor">Background color for matching cell</param>
    /// <param name="fontColor">Font color for matching cell (optional)</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell HighlightIfGreaterThan(double value, XLColor backgroundColor, XLColor? fontColor = null)
    {
        var range = _cell.AsRange();
        var conditionalFormat = range.AddConditionalFormat();
        var format = conditionalFormat.WhenGreaterThan(value);
        format.Fill.SetBackgroundColor(backgroundColor);
        
        if (fontColor != null)
        {
            format.Font.SetFontColor(fontColor);
        }
        
        return this;
    }

    /// <summary>
    /// Highlights the cell if it is less than a specific value
    /// </summary>
    /// <param name="value">Value to compare against</param>
    /// <param name="backgroundColor">Background color for matching cell</param>
    /// <param name="fontColor">Font color for matching cell (optional)</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell HighlightIfLessThan(double value, XLColor backgroundColor, XLColor? fontColor = null)
    {
        var range = _cell.AsRange();
        var conditionalFormat = range.AddConditionalFormat();
        var format = conditionalFormat.WhenLessThan(value);
        format.Fill.SetBackgroundColor(backgroundColor);
        
        if (fontColor != null)
        {
            format.Font.SetFontColor(fontColor);
        }
        
        return this;
    }

    /// <summary>
    /// Highlights the cell if it is between two values (inclusive)
    /// </summary>
    /// <param name="minValue">Minimum value</param>
    /// <param name="maxValue">Maximum value</param>
    /// <param name="backgroundColor">Background color for matching cell</param>
    /// <param name="fontColor">Font color for matching cell (optional)</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell HighlightIfBetween(double minValue, double maxValue, XLColor backgroundColor, XLColor? fontColor = null)
    {
        var range = _cell.AsRange();
        var conditionalFormat = range.AddConditionalFormat();
        var format = conditionalFormat.WhenBetween(minValue, maxValue);
        format.Fill.SetBackgroundColor(backgroundColor);
        
        if (fontColor != null)
        {
            format.Font.SetFontColor(fontColor);
        }
        
        return this;
    }

    /// <summary>
    /// Highlights the cell if it is not between two values
    /// </summary>
    /// <param name="minValue">Minimum value</param>
    /// <param name="maxValue">Maximum value</param>
    /// <param name="backgroundColor">Background color for matching cell</param>
    /// <param name="fontColor">Font color for matching cell (optional)</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell HighlightIfNotBetween(double minValue, double maxValue, XLColor backgroundColor, XLColor? fontColor = null)
    {
        var range = _cell.AsRange();
        var conditionalFormat = range.AddConditionalFormat();
        var format = conditionalFormat.WhenNotBetween(minValue, maxValue);
        format.Fill.SetBackgroundColor(backgroundColor);
        
        if (fontColor != null)
        {
            format.Font.SetFontColor(fontColor);
        }
        
        return this;
    }

    /// <summary>
    /// Highlights the cell if it is blank
    /// </summary>
    /// <param name="backgroundColor">Background color for blank cell</param>
    /// <param name="fontColor">Font color for blank cell (optional)</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell HighlightIfBlank(XLColor backgroundColor, XLColor? fontColor = null)
    {
        var range = _cell.AsRange();
        var conditionalFormat = range.AddConditionalFormat();
        var format = conditionalFormat.WhenIsBlank();
        format.Fill.SetBackgroundColor(backgroundColor);
        
        if (fontColor != null)
        {
            format.Font.SetFontColor(fontColor);
        }
        
        return this;
    }

    /// <summary>
    /// Highlights the cell if it contains an error
    /// </summary>
    /// <param name="backgroundColor">Background color for error cell</param>
    /// <param name="fontColor">Font color for error cell (optional)</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell HighlightIfError(XLColor backgroundColor, XLColor? fontColor = null)
    {
        var range = _cell.AsRange();
        var conditionalFormat = range.AddConditionalFormat();
        var format = conditionalFormat.WhenIsError();
        format.Fill.SetBackgroundColor(backgroundColor);
        
        if (fontColor != null)
        {
            format.Font.SetFontColor(fontColor);
        }
        
        return this;
    }

    /// <summary>
    /// Applies a formula-based conditional formatting rule to the cell
    /// </summary>
    /// <param name="formula">Formula to evaluate (without = sign)</param>
    /// <param name="backgroundColor">Background color when formula is true</param>
    /// <param name="fontColor">Font color when formula is true (optional)</param>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell HighlightWhenFormula(string formula, XLColor backgroundColor, XLColor? fontColor = null)
    {
        var range = _cell.AsRange();
        var conditionalFormat = range.AddConditionalFormat();
        var format = conditionalFormat.WhenIsTrue(formula);
        format.Fill.SetBackgroundColor(backgroundColor);
        
        if (fontColor != null)
        {
            format.Font.SetFontColor(fontColor);
        }
        
        return this;
    }

    /// <summary>
    /// Clears all conditional formatting from the cell
    /// </summary>
    /// <returns>This FluentCell for method chaining</returns>
    public FluentCell ClearConditionalFormatting()
    {
        var range = _cell.AsRange();
        var worksheet = _cell.Worksheet;
        var formatsToRemove = new List<IXLConditionalFormat>();
        
        foreach (var cf in worksheet.ConditionalFormats.ToList())
        {
            if (cf.Ranges.Contains(range))
            {
                formatsToRemove.Add(cf);
            }
        }
        
        // Remove them individually
        foreach (var cf in formatsToRemove)
        {
            worksheet.ConditionalFormats.Remove(cf => true);
        }
        
        return this;
    }

    #endregion
}