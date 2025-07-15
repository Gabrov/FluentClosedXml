using ClosedXML.Excel;

namespace FluentClosedXml.Extensions;

/// <summary>
/// Extension methods for enhanced fluent API functionality
/// </summary>
public static class FluentExtensions
{
    /// <summary>
    /// Converts an enumerable of objects to a fluent workbook with data
    /// </summary>
    /// <typeparam name="T">Type of objects in the enumerable</typeparam>
    /// <param name="data">Data to convert</param>
    /// <param name="worksheetName">Name of the worksheet</param>
    /// <param name="createTable">Whether to create an Excel table</param>
    /// <returns>FluentWorkbook containing the data</returns>
    public static FluentWorkbook ToFluentWorkbook<T>(this IEnumerable<T> data, 
        string worksheetName = "Data", bool createTable = true)
    {
        var workbook = FluentWorkbook.Create();
        var worksheet = workbook.AddWorksheet(worksheetName);
        worksheet.AddData(data, createTable: createTable);
        
        if (createTable)
        {
            worksheet.AutoFitColumns();
        }
        
        return workbook;
    }

    /// <summary>
    /// Creates a quick report with headers and data
    /// </summary>
    /// <typeparam name="T">Type of objects in the data</typeparam>
    /// <param name="data">Data for the report</param>
    /// <param name="title">Report title</param>
    /// <param name="worksheetName">Worksheet name</param>
    /// <returns>FluentWorkbook with formatted report</returns>
    public static FluentWorkbook ToReport<T>(this IEnumerable<T> data, 
        string title = "Report", string worksheetName = "Report")
    {
        var workbook = FluentWorkbook.Create();
        var worksheet = workbook.AddWorksheet(worksheetName);
        
        // Add title
        worksheet.SetCell("A1", title)
                .Bold()
                .WithFontSize(16)
                .WithFontColor(XLColor.DarkBlue);
        
        // Add data starting from row 3
        var dataRange = worksheet.AddData(data, "A3", createTable: true);
        
        // Format the report
        worksheet.AutoFitColumns()
                .FreezeTopRow();
        
        return workbook;
    }

    /// <summary>
    /// Applies a predefined style theme to a range
    /// </summary>
    /// <param name="range">Range to style</param>
    /// <param name="theme">Theme to apply</param>
    /// <returns>The styled FluentRange</returns>
    public static FluentRange WithTheme(this FluentRange range, FluentTheme theme)
    {
        return theme switch
        {
            FluentTheme.Header => range.Bold()
                                      .WithBackgroundColor(XLColor.LightBlue)
                                      .WithFontColor(XLColor.White)
                                      .Center(),
            
            FluentTheme.Data => range.WithBorder()
                                    .WithBackgroundColor(XLColor.White),
            
            FluentTheme.Total => range.Bold()
                                     .WithBackgroundColor(XLColor.LightGray)
                                     .WithBorder(XLBorderStyleValues.Thick),
            
            FluentTheme.Warning => range.WithBackgroundColor(XLColor.Yellow)
                                       .WithFontColor(XLColor.DarkRed)
                                       .Bold(),
            
            FluentTheme.Success => range.WithBackgroundColor(XLColor.LightGreen)
                                       .WithFontColor(XLColor.DarkGreen)
                                       .Bold(),
            
            _ => range
        };
    }

    /// <summary>
    /// Applies a predefined style theme to a cell
    /// </summary>
    /// <param name="cell">Cell to style</param>
    /// <param name="theme">Theme to apply</param>
    /// <returns>The styled FluentCell</returns>
    public static FluentCell WithTheme(this FluentCell cell, FluentTheme theme)
    {
        return theme switch
        {
            FluentTheme.Header => cell.Bold()
                                     .WithBackgroundColor(XLColor.LightBlue)
                                     .WithFontColor(XLColor.White)
                                     .Center(),
            
            FluentTheme.Data => cell.WithBorder()
                                   .WithBackgroundColor(XLColor.White),
            
            FluentTheme.Total => cell.Bold()
                                    .WithBackgroundColor(XLColor.LightGray)
                                    .WithBorder(XLBorderStyleValues.Thick),
            
            FluentTheme.Warning => cell.WithBackgroundColor(XLColor.Yellow)
                                      .WithFontColor(XLColor.DarkRed)
                                      .Bold(),
            
            FluentTheme.Success => cell.WithBackgroundColor(XLColor.LightGreen)
                                      .WithFontColor(XLColor.DarkGreen)
                                      .Bold(),
            
            _ => cell
        };
    }

    /// <summary>
    /// Creates a formatted financial report
    /// </summary>
    /// <param name="data">Financial data</param>
    /// <param name="title">Report title</param>
    /// <returns>FluentWorkbook with financial report</returns>
    public static FluentWorkbook ToFinancialReport<T>(this IEnumerable<T> data, string title = "Financial Report") 
        where T : class
    {
        var workbook = FluentWorkbook.Create();
        var worksheet = workbook.AddWorksheet("Financial Report");
        
        // Title
        worksheet.SetCell("A1", title)
                .WithTheme(FluentTheme.Header)
                .WithFontSize(18);
        
        // Date
        worksheet.SetCell("A2", $"Generated: {DateTime.Now:yyyy-MM-dd HH:mm}")
                .Italic()
                .WithFontSize(10);
        
        // Data
        var dataRange = worksheet.AddData(data, "A4", createTable: true);
        
        // Apply financial formatting to numeric columns
        var properties = typeof(T).GetProperties()
            .Where(p => p.PropertyType == typeof(decimal) || p.PropertyType == typeof(double) || 
                       p.PropertyType == typeof(decimal?) || p.PropertyType == typeof(double?))
            .ToList();
        
        for (int i = 0; i < properties.Count; i++)
        {
            var columnIndex = i + 1; // Assuming properties match column order
            var columnRange = worksheet.GetRange(5, columnIndex, dataRange.Range.LastRow().RowNumber(), columnIndex);
            columnRange.AsCurrency();
        }
        
        worksheet.AutoFitColumns()
                .FreezeTopRow();
        
        return workbook;
    }
}

/// <summary>
/// Predefined themes for styling
/// </summary>
public enum FluentTheme
{
    Header,
    Data,
    Total,
    Warning,
    Success
}