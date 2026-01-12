using System;
using System.Linq;
using System.Collections.Generic;

using ClosedXML.Excel;

namespace FluentClosedXml
{

    /// <summary>
    /// Fluent API wrapper for ClosedXML range operations
    /// </summary>
    public class FluentRange
    {
        private readonly IXLRange _range;

        internal FluentRange(IXLRange range)
        {
            _range = range ?? throw new ArgumentNullException(nameof(range));
        }

        /// <summary>
        /// Gets the underlying ClosedXML range for advanced operations
        /// </summary>
        public IXLRange Range => _range;

        /// <summary>
        /// Gets the range address as a string (e.g., "A1:B10")
        /// </summary>
        /// <returns>The range address in A1 notation</returns>
        public string GetAddress()
        {
            return _range.RangeAddress.ToString();
        }

        /// <summary>
        /// Sets values for a range with automatic orientation detection
        /// </summary>
        /// <param name="values">array of values</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithValues(object[] values)
        {
            // Determine if the range is horizontal or vertical based on its dimensions
            var rowCount = _range.RowCount();
            var colCount = _range.ColumnCount();

            // If range has only one row, treat as horizontal; if only one column, treat as vertical
            // If it's a square or rectangular range, default to vertical (existing behavior)
            var isHorizontal = rowCount == 1 && colCount > 1;

            return WithValues(values, isHorizontal ? RangeOrientation.Horizontal : RangeOrientation.Vertical);
        }

        /// <summary>
        /// Sets values for a range with specified orientation
        /// </summary>
        /// <param name="values">array of values</param>
        /// <param name="orientation">Orientation for placing values (horizontal or vertical)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithValues(object[] values, RangeOrientation orientation)
        {
            var valueCount = values.Length;

            if (orientation == RangeOrientation.Horizontal)
            {
                // Place values horizontally across columns in the first row of the range
                for (int i = 0; i < valueCount; i++)
                {
                    var cell = _range.Cell(1, i + 1); // First row, columns 1, 2, 3, etc.
                    var value = values[i];
                    if (value != null)
                    {
                        cell.Value = XLCellValue.FromObject(value);
                    }
                }
            }
            else
            {
                // Place values vertically down rows in the first column of the range (existing behavior)
                for (int i = 0; i < valueCount; i++)
                {
                    var cell = _range.Cell(i + 1, 1); // Rows 1, 2, 3, etc., first column
                    var value = values[i];
                    if (value != null)
                    {
                        cell.Value = XLCellValue.FromObject(value);
                    }
                }
            }

            return this;
        }

        /// <summary>
        /// Sets values for the entire range
        /// </summary>
        /// <param name="values">2D array of values</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithValues(object[,] values)
        {
            var rowCount = values.GetLength(0);
            var colCount = values.GetLength(1);

            for (int row = 0; row < rowCount; row++)
            {
                for (int col = 0; col < colCount; col++)
                {
                    var cell = _range.Cell(row + 1, col + 1);
                    var value = values[row, col];
                    if (value != null)
                    {
                        cell.Value = XLCellValue.FromObject(value);
                    }
                }
            }

            return this;
        }

        /// <summary>
        /// Sets a single value for the entire range
        /// </summary>
        /// <param name="value">Value to set</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithValue(object value)
        {
            if (value != null)
            {
                _range.Value = XLCellValue.FromObject(value);
            }
            return this;
        }

        /// <summary>
        /// Sets a formula for the range using A1 notation
        /// </summary>
        /// <param name="formula">Formula to set (without = sign)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithFormula(string formula)
        {
            // Set the formula on the first cell of the range (which becomes the merged cell)
            _range.FirstCell().FormulaA1 = formula;
            return this;
        }

        /// <summary>
        /// Sets a formula for the range using R1C1 notation
        /// </summary>
        /// <param name="formula">Formula in R1C1 notation (without = sign)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithFormulaR1C1(string formula)
        {
            // Set the formula on the first cell of the range (which becomes the merged cell)
            _range.FirstCell().FormulaR1C1 = formula;
            return this;
        }

        #region Common Formula Methods

        /// <summary>
        /// Sets a SUM formula for the specified range
        /// </summary>
        /// <param name="range">Range to sum (e.g., "A1:A10")</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithSum(string range)
        {
            _range.FirstCell().FormulaA1 = $"SUM({range})";
            return this;
        }

        /// <summary>
        /// Sets a SUM formula for multiple ranges with optional negative signs
        /// </summary>
        /// <param name="ranges">Multiple ranges to sum, with optional negative signs (e.g., "A1:A5", "-B21", "C1:C10")</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithSum(params string[] ranges)
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
            _range.FirstCell().FormulaA1 = formula;
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

        #endregion

        /// <summary>
        /// Sets the range style using a fluent interface
        /// </summary>
        /// <param name="styleAction">Action to configure range style</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithStyle(Action<IXLStyle> styleAction)
        {
            styleAction(_range.Style);
            return this;
        }

        /// <summary>
        /// Makes the range bold
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange Bold()
        {
            _range.Style.Font.Bold = true;
            return this;
        }

        /// <summary>
        /// Makes the range italic
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange Italic()
        {
            _range.Style.Font.Italic = true;
            return this;
        }

        /// <summary>
        /// Sets the font size for the range
        /// </summary>
        /// <param name="size">Font size</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithFontSize(double size)
        {
            _range.Style.Font.FontSize = size;
            return this;
        }

        /// <summary>
        /// Sets the font color for the range
        /// </summary>
        /// <param name="color">Font color</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithFontColor(XLColor color)
        {
            _range.Style.Font.FontColor = color;
            return this;
        }

        /// <summary>
        /// Sets the background color for the range
        /// </summary>
        /// <param name="color">Background color</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithBackgroundColor(XLColor color)
        {
            _range.Style.Fill.BackgroundColor = color;
            return this;
        }

        /// <summary>
        /// Sets horizontal alignment for the range
        /// </summary>
        /// <param name="alignment">Horizontal alignment</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithHorizontalAlignment(XLAlignmentHorizontalValues alignment)
        {
            _range.Style.Alignment.Horizontal = alignment;
            return this;
        }

        /// <summary>
        /// Sets vertical alignment for the range
        /// </summary>
        /// <param name="alignment">Vertical alignment</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithVerticalAlignment(XLAlignmentVerticalValues alignment)
        {
            _range.Style.Alignment.Vertical = alignment;
            return this;
        }

        /// <summary>
        /// Centers the range content both horizontally and vertically
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange Center()
        {
            _range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            _range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            return this;
        }

        /// <summary>
        /// Enables text wrapping for the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithTextWrap()
        {
            _range.Style.Alignment.WrapText = true;
            return this;
        }

        /// <summary>
        /// Disables text wrapping for the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithoutTextWrap()
        {
            _range.Style.Alignment.WrapText = false;
            return this;
        }

        /// <summary>
        /// Sets text wrapping for the range
        /// </summary>
        /// <param name="wrapText">True to enable text wrapping, false to disable</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithTextWrap(bool wrapText)
        {
            _range.Style.Alignment.WrapText = wrapText;
            return this;
        }

        /// <summary>
        /// Sets borders around the range
        /// </summary>
        /// <param name="borderStyle">Border style</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithBorder(XLBorderStyleValues borderStyle = XLBorderStyleValues.Thin)
        {
            _range.Style.Border.OutsideBorder = borderStyle;
            return this;
        }

        /// <summary>
        /// Sets inside borders for the range
        /// </summary>
        /// <param name="borderStyle">Border style</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithInsideBorder(XLBorderStyleValues borderStyle = XLBorderStyleValues.Thin)
        {
            _range.Style.Border.InsideBorder = borderStyle;
            return this;
        }

        /// <summary>
        /// Sets a thick border around the range with thin borders inside (common table formatting)
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithTableBorders()
        {
            return WithTableBorders(XLBorderStyleValues.Thick, XLBorderStyleValues.Thin);
        }

        /// <summary>
        /// Sets custom border styles around and inside the range
        /// </summary>
        /// <param name="outsideBorderStyle">Border style for the outside border</param>
        /// <param name="insideBorderStyle">Border style for the inside borders</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithTableBorders(XLBorderStyleValues outsideBorderStyle, XLBorderStyleValues insideBorderStyle)
        {
            _range.Style.Border.OutsideBorder = outsideBorderStyle;
            _range.Style.Border.InsideBorder = insideBorderStyle;
            return this;
        }

        /// <summary>
        /// Sets specific border sides with custom styles
        /// </summary>
        /// <param name="top">Top border style (null to leave unchanged)</param>
        /// <param name="right">Right border style (null to leave unchanged)</param>
        /// <param name="bottom">Bottom border style (null to leave unchanged)</param>
        /// <param name="left">Left border style (null to leave unchanged)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithBorders(XLBorderStyleValues? top = null, XLBorderStyleValues? right = null,
                                       XLBorderStyleValues? bottom = null, XLBorderStyleValues? left = null)
        {
            if (top.HasValue) _range.Style.Border.TopBorder = top.Value;
            if (right.HasValue) _range.Style.Border.RightBorder = right.Value;
            if (bottom.HasValue) _range.Style.Border.BottomBorder = bottom.Value;
            if (left.HasValue) _range.Style.Border.LeftBorder = left.Value;
            return this;
        }

        /// <summary>
        /// Sets border colors for the range
        /// </summary>
        /// <param name="color">Border color to apply to all borders</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithBorderColor(XLColor color)
        {
            _range.Style.Border.TopBorderColor = color;
            _range.Style.Border.RightBorderColor = color;
            _range.Style.Border.BottomBorderColor = color;
            _range.Style.Border.LeftBorderColor = color;
            return this;
        }

        /// <summary>
        /// Removes all borders from the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithoutBorder()
        {
            _range.Style.Border.OutsideBorder = XLBorderStyleValues.None;
            _range.Style.Border.InsideBorder = XLBorderStyleValues.None;
            return this;
        }

        /// <summary>
        /// Sets a thick border around the range only (no inside borders)
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithThickBorder()
        {
            return WithBorder(XLBorderStyleValues.Thick);
        }

        /// <summary>
        /// Sets a medium border around the range only (no inside borders)
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithMediumBorder()
        {
            return WithBorder(XLBorderStyleValues.Medium);
        }

        /// <summary>
        /// Sets a double border around the range only (no inside borders)
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithDoubleBorder()
        {
            return WithBorder(XLBorderStyleValues.Double);
        }

        /// <summary>
        /// Sets a dashed border around the range only (no inside borders)
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithDashedBorder()
        {
            return WithBorder(XLBorderStyleValues.Dashed);
        }

        /// <summary>
        /// Sets a dotted border around the range only (no inside borders)
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithDottedBorder()
        {
            return WithBorder(XLBorderStyleValues.Dotted);
        }

        /// <summary>
        /// Sets the number format for the range
        /// </summary>
        /// <param name="format">Number format string</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithNumberFormat(string format)
        {
            _range.Style.NumberFormat.Format = format;
            return this;
        }

        /// <summary>
        /// Formats the range as currency
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsCurrency()
        {
            _range.Style.NumberFormat.Format = "\"$\"#,##0.00_);(\"$\"#,##0.00)";
            return this;
        }

        /// <summary>
        /// Formats the range as percentage
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsPercentage()
        {
            _range.Style.NumberFormat.Format = "0.00%";
            return this;
        }

        /// <summary>
        /// Formats the range as date
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsDate()
        {
            _range.Style.NumberFormat.Format = "mm/dd/yyyy";
            return this;
        }

        #region Predefined Number Format Methods

        /// <summary>
        /// Sets the number format using ClosedXML's predefined format
        /// </summary>
        /// <param name="formatId">Predefined format ID</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithPredefinedFormat(int formatId)
        {
            _range.Style.NumberFormat.NumberFormatId = formatId;
            return this;
        }

        /// <summary>
        /// Applies General format (Excel's default) to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsGeneral()
        {
            _range.Style.NumberFormat.NumberFormatId = 0;
            return this;
        }

        /// <summary>
        /// Applies format: 0 (Number with no decimals) to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsNumber()
        {
            _range.Style.NumberFormat.NumberFormatId = 1;
            return this;
        }

        /// <summary>
        /// Applies format: 0.00 (Number with 2 decimals) to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsNumberWith2Decimals()
        {
            _range.Style.NumberFormat.NumberFormatId = 2;
            return this;
        }

        /// <summary>
        /// Applies format: #,##0 (Number with thousands separator) to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsNumberWithThousandsSeparator()
        {
            _range.Style.NumberFormat.NumberFormatId = 3;
            return this;
        }

        /// <summary>
        /// Applies format: #,##0.00 (Number with thousands separator and 2 decimals) to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsNumberWithThousandsSeparatorAnd2Decimals()
        {
            _range.Style.NumberFormat.NumberFormatId = 4;
            return this;
        }

        /// <summary>
        /// Applies standard currency format to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsCurrencyBuiltIn()
        {
            _range.Style.NumberFormat.NumberFormatId = 5;
            return this;
        }

        /// <summary>
        /// Applies currency format with red negative values to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsCurrencyWithRedNegatives()
        {
            _range.Style.NumberFormat.NumberFormatId = 6;
            return this;
        }

        /// <summary>
        /// Applies currency format with 2 decimals to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsCurrencyWith2Decimals()
        {
            _range.Style.NumberFormat.NumberFormatId = 7;
            return this;
        }

        /// <summary>
        /// Applies currency format with 2 decimals and red negative values to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsCurrencyWith2DecimalsAndRedNegatives()
        {
            _range.Style.NumberFormat.NumberFormatId = 8;
            return this;
        }

        /// <summary>
        /// Applies percentage format: 0% to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsPercentageBuiltIn()
        {
            _range.Style.NumberFormat.NumberFormatId = 9;
            return this;
        }

        /// <summary>
        /// Applies percentage format with 2 decimals: 0.00% to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsPercentageWith2Decimals()
        {
            _range.Style.NumberFormat.NumberFormatId = 10;
            return this;
        }

        /// <summary>
        /// Applies scientific notation format: 0.00E+00 to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsScientific()
        {
            _range.Style.NumberFormat.NumberFormatId = 11;
            return this;
        }

        /// <summary>
        /// Applies fraction format: # ?/? to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsFraction()
        {
            _range.Style.NumberFormat.NumberFormatId = 12;
            return this;
        }

        /// <summary>
        /// Applies fraction format with denominators up to 99: # ??/?? to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsFractionUpTo99()
        {
            _range.Style.NumberFormat.NumberFormatId = 13;
            return this;
        }

        /// <summary>
        /// Applies short date format: mm/dd/yyyy to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsShortDate()
        {
            _range.Style.NumberFormat.NumberFormatId = 14;
            return this;
        }

        /// <summary>
        /// Applies date format: d-mmm-yy to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsDateDMmmYy()
        {
            _range.Style.NumberFormat.NumberFormatId = 15;
            return this;
        }

        /// <summary>
        /// Applies date format: d-mmm to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsDateDMmm()
        {
            _range.Style.NumberFormat.NumberFormatId = 16;
            return this;
        }

        /// <summary>
        /// Applies date format: mmm-yy to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsDateMmmYy()
        {
            _range.Style.NumberFormat.NumberFormatId = 17;
            return this;
        }

        /// <summary>
        /// Applies time format: h:mm AM/PM to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsTimeAmPm()
        {
            _range.Style.NumberFormat.NumberFormatId = 18;
            return this;
        }

        /// <summary>
        /// Applies time format with seconds: h:mm:ss AM/PM to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsTimeWithSecondsAmPm()
        {
            _range.Style.NumberFormat.NumberFormatId = 19;
            return this;
        }

        /// <summary>
        /// Applies 24-hour time format: h:mm to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsTime24Hour()
        {
            _range.Style.NumberFormat.NumberFormatId = 20;
            return this;
        }

        /// <summary>
        /// Applies 24-hour time format with seconds: h:mm:ss to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsTime24HourWithSeconds()
        {
            _range.Style.NumberFormat.NumberFormatId = 21;
            return this;
        }

        /// <summary>
        /// Applies date and time format: mm/dd/yyyy h:mm to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsDateTime()
        {
            _range.Style.NumberFormat.NumberFormatId = 22;
            return this;
        }

        /// <summary>
        /// Applies accounting format with no decimals to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsAccounting()
        {
            _range.Style.NumberFormat.NumberFormatId = 37;
            return this;
        }

        /// <summary>
        /// Applies accounting format with 2 decimals to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsAccountingWith2Decimals()
        {
            _range.Style.NumberFormat.NumberFormatId = 38;
            return this;
        }

        /// <summary>
        /// Applies accounting format with red negatives to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsAccountingWithRedNegatives()
        {
            _range.Style.NumberFormat.NumberFormatId = 39;
            return this;
        }

        /// <summary>
        /// Applies accounting format with 2 decimals and red negatives to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsAccountingWith2DecimalsAndRedNegatives()
        {
            _range.Style.NumberFormat.NumberFormatId = 40;
            return this;
        }

        /// <summary>
        /// Applies text format (displays numbers as text) to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsText()
        {
            _range.Style.NumberFormat.NumberFormatId = 49;
            return this;
        }

        /// <summary>
        /// Sets format for elapsed time in hours: [h]:mm:ss to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsElapsedTimeHours()
        {
            _range.Style.NumberFormat.Format = "[h]:mm:ss";
            return this;
        }

        /// <summary>
        /// Sets format for elapsed time in minutes: [mm]:ss to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsElapsedTimeMinutes()
        {
            _range.Style.NumberFormat.Format = "[mm]:ss";
            return this;
        }

        /// <summary>
        /// Sets format for elapsed time in seconds: [ss] to the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsElapsedTimeSeconds()
        {
            _range.Style.NumberFormat.Format = "[ss]";
            return this;
        }

        /// <summary>
        /// Applies number format with specific decimal places to the range
        /// </summary>
        /// <param name="decimalPlaces">Number of decimal places (0-30)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsNumberWithDecimals(int decimalPlaces)
        {
            if (decimalPlaces < 0 || decimalPlaces > 30)
                throw new ArgumentOutOfRangeException(nameof(decimalPlaces), "Decimal places must be between 0 and 30");

            var format = decimalPlaces == 0 ? "0" : "0." + new string('0', decimalPlaces);
            _range.Style.NumberFormat.Format = format;
            return this;
        }

        /// <summary>
        /// Applies number format with thousands separator and specific decimal places to the range
        /// </summary>
        /// <param name="decimalPlaces">Number of decimal places (0-30)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsNumberWithThousandsSeparatorAndDecimals(int decimalPlaces)
        {
            if (decimalPlaces < 0 || decimalPlaces > 30)
                throw new ArgumentOutOfRangeException(nameof(decimalPlaces), "Decimal places must be between 0 and 30");

            var format = decimalPlaces == 0 ? "#,##0" : "#,##0." + new string('0', decimalPlaces);
            _range.Style.NumberFormat.Format = format;
            return this;
        }

        /// <summary>
        /// Applies percentage format with specific decimal places to the range
        /// </summary>
        /// <param name="decimalPlaces">Number of decimal places (0-30)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsPercentageWithDecimals(int decimalPlaces)
        {
            if (decimalPlaces < 0 || decimalPlaces > 30)
                throw new ArgumentOutOfRangeException(nameof(decimalPlaces), "Decimal places must be between 0 and 30");

            var format = decimalPlaces == 0 ? "0%" : "0." + new string('0', decimalPlaces) + "%";
            _range.Style.NumberFormat.Format = format;
            return this;
        }

        /// <summary>
        /// Applies currency format with specific decimal places and custom symbol to the range
        /// </summary>
        /// <param name="decimalPlaces">Number of decimal places (0-30)</param>
        /// <param name="currencySymbol">Currency symbol (default: $)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsCurrencyWithDecimals(int decimalPlaces, string currencySymbol = "$")
        {
            if (decimalPlaces < 0 || decimalPlaces > 30)
                throw new ArgumentOutOfRangeException(nameof(decimalPlaces), "Decimal places must be between 0 and 30");

            var decimalsFormat = decimalPlaces == 0 ? "" : "." + new string('0', decimalPlaces);
            var format = $"\"{currencySymbol}\"#,##0{decimalsFormat}_);(\"{currencySymbol}\"#,##0{decimalsFormat})";
            _range.Style.NumberFormat.Format = format;
            return this;
        }

        /// <summary>
        /// Gets the current number format ID of the range
        /// </summary>
        /// <returns>The number format ID</returns>
        public int GetNumberFormatId()
        {
            return _range.Style.NumberFormat.NumberFormatId;
        }

        /// <summary>
        /// Gets the current number format string of the range
        /// </summary>
        /// <returns>The number format string</returns>
        public string GetNumberFormat()
        {
            return _range.Style.NumberFormat.Format;
        }

        #endregion

        /// <summary>
        /// Auto-fits columns in the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AutoFitColumns()
        {
            var worksheet = _range.Worksheet;
            var firstColumn = _range.FirstColumn().ColumnNumber();
            var lastColumn = _range.LastColumn().ColumnNumber();

            for (int col = firstColumn; col <= lastColumn; col++)
            {
                worksheet.Column(col).AdjustToContents();
            }
            return this;
        }

        /// <summary>
        /// Auto-fits rows in the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AutoFitRows()
        {
            var worksheet = _range.Worksheet;
            var firstRow = _range.FirstRow().RowNumber();
            var lastRow = _range.LastRow().RowNumber();

            for (int row = firstRow; row <= lastRow; row++)
            {
                worksheet.Row(row).AdjustToContents();
            }
            return this;
        }

        /// <summary>
        /// Creates an Excel table from the range
        /// </summary>
        /// <param name="tableName">Optional table name</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange AsTable(string tableName = null)
        {
            var table = _range.CreateTable();
            if (!string.IsNullOrEmpty(tableName))
            {
                table.Name = tableName;
            }
            return this;
        }

        /// <summary>
        /// Merges the cells in the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange Merge()
        {
            _range.Merge();
            return this;
        }

        /// <summary>
        /// Unmerges the cells in the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange Unmerge()
        {
            _range.Unmerge();
            return this;
        }

        /// <summary>
        /// Clears the content of the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange Clear()
        {
            _range.Clear();
            return this;
        }

        /// <summary>
        /// Gets a specific cell from the range
        /// </summary>
        /// <param name="row">Row offset (1-based)</param>
        /// <param name="column">Column offset (1-based)</param>
        /// <returns>FluentCell for cell operations</returns>
        public FluentCell GetCell(int row, int column)
        {
            var cell = _range.Cell(row, column);
            return new FluentCell(cell);
        }

        /// <summary>
        /// Applies conditional formatting to the range
        /// </summary>
        /// <param name="configure">Action to configure conditional formatting</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithConditionalFormatting(Action<IXLConditionalFormat> configure)
        {
            var conditionalFormat = _range.AddConditionalFormat();
            configure(conditionalFormat);
            return this;
        }

        #region Conditional Formatting Methods

        /// <summary>
        /// Highlights cells that contain specific text
        /// </summary>
        /// <param name="text">Text to search for</param>
        /// <param name="backgroundColor">Background color for matching cells</param>
        /// <param name="fontColor">Font color for matching cells (optional)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange HighlightCellsContaining(string text, XLColor backgroundColor, XLColor fontColor = null)
        {
            var conditionalFormat = _range.AddConditionalFormat();
            var format = conditionalFormat.WhenContains(text);
            format.Fill.SetBackgroundColor(backgroundColor);

            if (fontColor != null)
            {
                format.Font.SetFontColor(fontColor);
            }

            return this;
        }

        /// <summary>
        /// Highlights cells that are greater than a specific value
        /// </summary>
        /// <param name="value">Value to compare against</param>
        /// <param name="backgroundColor">Background color for matching cells</param>
        /// <param name="fontColor">Font color for matching cells (optional)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange HighlightCellsGreaterThan(double value, XLColor backgroundColor, XLColor fontColor = null)
        {
            var conditionalFormat = _range.AddConditionalFormat();
            var format = conditionalFormat.WhenGreaterThan(value);
            format.Fill.SetBackgroundColor(backgroundColor);

            if (fontColor != null)
            {
                format.Font.SetFontColor(fontColor);
            }

            return this;
        }

        /// <summary>
        /// Highlights cells that are less than a specific value
        /// </summary>
        /// <param name="value">Value to compare against</param>
        /// <param name="backgroundColor">Background color for matching cells</param>
        /// <param name="fontColor">Font color for matching cells (optional)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange HighlightCellsLessThan(double value, XLColor backgroundColor, XLColor fontColor = null)
        {
            var conditionalFormat = _range.AddConditionalFormat();
            var format = conditionalFormat.WhenLessThan(value);
            format.Fill.SetBackgroundColor(backgroundColor);

            if (fontColor != null)
            {
                format.Font.SetFontColor(fontColor);
            }

            return this;
        }

        /// <summary>
        /// Highlights cells that are between two values (inclusive)
        /// </summary>
        /// <param name="minValue">Minimum value</param>
        /// <param name="maxValue">Maximum value</param>
        /// <param name="backgroundColor">Background color for matching cells</param>
        /// <param name="fontColor">Font color for matching cells (optional)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange HighlightCellsBetween(double minValue, double maxValue, XLColor backgroundColor, XLColor fontColor = null)
        {
            var conditionalFormat = _range.AddConditionalFormat();
            var format = conditionalFormat.WhenBetween(minValue, maxValue);
            format.Fill.SetBackgroundColor(backgroundColor);

            if (fontColor != null)
            {
                format.Font.SetFontColor(fontColor);
            }

            return this;
        }

        /// <summary>
        /// Highlights cells that are not between two values
        /// </summary>
        /// <param name="minValue">Minimum value</param>
        /// <param name="maxValue">Maximum value</param>
        /// <param name="backgroundColor">Background color for matching cells</param>
        /// <param name="fontColor">Font color for matching cells (optional)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange HighlightCellsNotBetween(double minValue, double maxValue, XLColor backgroundColor, XLColor fontColor = null)
        {
            var conditionalFormat = _range.AddConditionalFormat();
            var format = conditionalFormat.WhenNotBetween(minValue, maxValue);
            format.Fill.SetBackgroundColor(backgroundColor);

            if (fontColor != null)
            {
                format.Font.SetFontColor(fontColor);
            }

            return this;
        }

        /// <summary>
        /// Highlights blank cells in the range
        /// </summary>
        /// <param name="backgroundColor">Background color for blank cells</param>
        /// <param name="fontColor">Font color for blank cells (optional)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange HighlightBlankCells(XLColor backgroundColor, XLColor fontColor = null)
        {
            var conditionalFormat = _range.AddConditionalFormat();
            var format = conditionalFormat.WhenIsBlank();
            format.Fill.SetBackgroundColor(backgroundColor);

            if (fontColor != null)
            {
                format.Font.SetFontColor(fontColor);
            }

            return this;
        }

        /// <summary>
        /// Highlights cells containing errors
        /// </summary>
        /// <param name="backgroundColor">Background color for error cells</param>
        /// <param name="fontColor">Font color for error cells (optional)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange HighlightErrorCells(XLColor backgroundColor, XLColor fontColor = null)
        {
            var conditionalFormat = _range.AddConditionalFormat();
            var format = conditionalFormat.WhenIsError();
            format.Fill.SetBackgroundColor(backgroundColor);

            if (fontColor != null)
            {
                format.Font.SetFontColor(fontColor);
            }

            return this;
        }

        /// <summary>
        /// Highlights top N values in the range
        /// </summary>
        /// <param name="n">Number of top values to highlight</param>
        /// <param name="backgroundColor">Background color for top values</param>
        /// <param name="fontColor">Font color for top values (optional)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange HighlightTopValues(int n, XLColor backgroundColor, XLColor fontColor = null)
        {
            var conditionalFormat = _range.AddConditionalFormat();
            var format = conditionalFormat.WhenIsTop(n, XLTopBottomType.Items);
            format.Fill.SetBackgroundColor(backgroundColor);

            if (fontColor != null)
            {
                format.Font.SetFontColor(fontColor);
            }

            return this;
        }

        /// <summary>
        /// Highlights bottom N values in the range
        /// </summary>
        /// <param name="n">Number of bottom values to highlight</param>
        /// <param name="backgroundColor">Background color for bottom values</param>
        /// <param name="fontColor">Font color for bottom values (optional)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange HighlightBottomValues(int n, XLColor backgroundColor, XLColor fontColor = null)
        {
            var conditionalFormat = _range.AddConditionalFormat();
            var format = conditionalFormat.WhenIsBottom(n, XLTopBottomType.Items);
            format.Fill.SetBackgroundColor(backgroundColor);

            if (fontColor != null)
            {
                format.Font.SetFontColor(fontColor);
            }

            return this;
        }

        /// <summary>
        /// Applies a formula-based conditional formatting rule
        /// </summary>
        /// <param name="formula">Formula to evaluate (without = sign)</param>
        /// <param name="backgroundColor">Background color when formula is true</param>
        /// <param name="fontColor">Font color when formula is true (optional)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange HighlightWhenFormula(string formula, XLColor backgroundColor, XLColor fontColor = null)
        {
            var conditionalFormat = _range.AddConditionalFormat();
            var format = conditionalFormat.WhenIsTrue(formula);
            format.Fill.SetBackgroundColor(backgroundColor);

            if (fontColor != null)
            {
                format.Font.SetFontColor(fontColor);
            }

            return this;
        }

        /// <summary>
        /// Applies a 3-color color scale to the range (red-yellow-green by default)
        /// </summary>
        /// <param name="minColor">Color for minimum values (default: Red)</param>
        /// <param name="midColor">Color for middle values (default: Yellow)</param>
        /// <param name="maxColor">Color for maximum values (default: Green)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithColorScale(XLColor minColor = null, XLColor midColor = null, XLColor maxColor = null)
        {
            var conditionalFormat = _range.AddConditionalFormat();
            conditionalFormat.ColorScale()
                .Minimum(XLCFContentType.Percent, "0", minColor ?? XLColor.Red)
                .Midpoint(XLCFContentType.Percent, "50", midColor ?? XLColor.Yellow)
                .Maximum(XLCFContentType.Percent, "100", maxColor ?? XLColor.Green);

            return this;
        }

        /// <summary>
        /// Applies a 2-color color scale to the range
        /// </summary>
        /// <param name="minColor">Color for minimum values (default: Red)</param>
        /// <param name="maxColor">Color for maximum values (default: Green)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithTwoColorScale(XLColor minColor = null, XLColor maxColor = null)
        {
            var conditionalFormat = _range.AddConditionalFormat();
            conditionalFormat.ColorScale()
                .Minimum(XLCFContentType.Percent, "0", minColor ?? XLColor.Red)
                .Maximum(XLCFContentType.Percent, "100", maxColor ?? XLColor.Green);

            return this;
        }

        /// <summary>
        /// Applies data bars to the range
        /// </summary>
        /// <param name="color">Color of the data bars (default: Blue)</param>
        /// <param name="showValue">Whether to show the cell value (default: true)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithDataBars(XLColor color = null, bool showValue = true)
        {
            var conditionalFormat = _range.AddConditionalFormat();
            conditionalFormat.DataBar(color ?? XLColor.Blue, showValue);

            return this;
        }

        /// <summary>
        /// Applies icon sets to the range
        /// </summary>
        /// <param name="iconSetType">Type of icon set to use (default: ThreeTrafficLights1)</param>
        /// <param name="showIconOnly">Whether to show only icons without values (default: false)</param>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange WithIconSet(XLIconSetStyle iconSetType = XLIconSetStyle.ThreeTrafficLights1, bool showIconOnly = false)
        {
            var conditionalFormat = _range.AddConditionalFormat();
            conditionalFormat.IconSet(iconSetType, showIconOnly);

            return this;
        }

        /// <summary>
        /// Clears all conditional formatting from the range
        /// </summary>
        /// <returns>This FluentRange for method chaining</returns>
        public FluentRange ClearConditionalFormatting()
        {
            // Find and remove conditional formats that apply to this range
            var worksheet = _range.Worksheet;
            var formatsToRemove = new List<IXLConditionalFormat>();

            foreach (var cf in worksheet.ConditionalFormats.ToList())
            {
                if (cf.Ranges.Contains(_range))
                {
                    formatsToRemove.Add(cf);
                }
            }

            // Remove them individually
            foreach (var cf in formatsToRemove)
            {
                worksheet.ConditionalFormats.Remove(lcf => true);
            }

            return this;
        }

        #endregion
    }

    /// <summary>
    /// Specifies the orientation for placing values in a range
    /// </summary>
    public enum RangeOrientation
    {
        /// <summary>
        /// Values are placed vertically (down rows)
        /// </summary>
        Vertical,

        /// <summary>
        /// Values are placed horizontally (across columns)
        /// </summary>
        Horizontal
    }
}
