using FluentClosedXml.Extensions;
using ClosedXML.Excel;

namespace FluentClosedXml.Demo;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("FluentClosedXml Demo - Creating Excel files with fluent API");
        Console.WriteLine("==========================================================");

        // Demo 1: Basic workbook and worksheet operations
        BasicDemo();

        // Demo 2: Working with data collections
        DataCollectionDemo();

        // Demo 3: Advanced styling and formatting
        AdvancedStylingDemo();

        // Demo 4: Financial report example
        FinancialReportDemo();

        // Demo 5: Extension methods demo
        ExtensionMethodsDemo();

        // Demo 6: Formula capabilities demo
        FormulaDemo();

        // Demo 7: Predefined number format demo
        PredefinedNumberFormatDemo();

        // Demo 8: Range address demo
        RangeAddressDemo();

        // Demo 9: Range formula and merge demo
        RangeFormulaAndMergeDemo();

        // Demo 10: Horizontal range values demo  
        HorizontalRangeValuesDemo();

        // Demo 11: Border demo
        BorderDemo();

        // Demo 12: Text wrap demo
        TextWrapDemo();

        // Demo 13: Multiple range SUM demo
        MultipleRangeSumDemo();

        Console.WriteLine("\nAll demos completed! Check the generated Excel files.");
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }

    static void BasicDemo()
    {
        Console.WriteLine("\n1. Basic Demo - Creating a simple workbook");

        using var workbook = FluentWorkbook.Create();
        
        var worksheet = workbook.AddWorksheet("Basic Demo");
        
        worksheet.SetCell("A1", "Hello, FluentClosedXml!")
            .Bold()
            .WithFontSize(16)
            .WithFontColor(XLColor.DarkBlue);
            
        worksheet.SetCell("A2", "This is a fluent API wrapper for ClosedXML")
            .Italic();
            
        worksheet.SetCell("A4", "Today's Date:")
            .Bold();
            
        worksheet.SetCell("B4", DateTime.Now)
            .AsDate();
            
        worksheet.SetCell("A5", "Random Number:")
            .Bold();
            
        worksheet.SetCell("B5", new Random().NextDouble() * 1000)
            .AsCurrency();
            
        worksheet.AutoFitColumns();

        workbook.SaveAs("basic-demo.xlsx");
        Console.WriteLine("   ✓ Created basic-demo.xlsx");
    }

    static void DataCollectionDemo()
    {
        Console.WriteLine("\n2. Data Collection Demo - Working with lists of objects");

        var employees = new[]
        {
            new Employee { Id = 1, Name = "John Doe", Department = "IT", Salary = 75000, HireDate = new DateTime(2020, 1, 15) },
            new Employee { Id = 2, Name = "Jane Smith", Department = "HR", Salary = 65000, HireDate = new DateTime(2019, 3, 22) },
            new Employee { Id = 3, Name = "Bob Johnson", Department = "Finance", Salary = 80000, HireDate = new DateTime(2021, 7, 10) },
            new Employee { Id = 4, Name = "Alice Brown", Department = "IT", Salary = 72000, HireDate = new DateTime(2020, 11, 5) },
            new Employee { Id = 5, Name = "Charlie Wilson", Department = "Marketing", Salary = 68000, HireDate = new DateTime(2022, 2, 28) }
        };

        using var workbook = FluentWorkbook.Create();
        
        var worksheet = workbook.AddWorksheet("Employee Data");

        // Add title
        worksheet.SetCell("A1", "Employee Directory")
            .Bold()
            .WithFontSize(18)
            .WithFontColor(XLColor.DarkBlue);

        // Add headers manually with styling
        var headers = new[] { "ID", "Name", "Department", "Salary", "Hire Date" };
        var headerRange = worksheet.AddHeaders(headers, "A3");
        headerRange.WithTheme(FluentTheme.Header);

        // Add employee data
        var row = 4;
        foreach (var emp in employees)
        {
            worksheet.SetCell(row, 1, emp.Id);
            worksheet.SetCell(row, 2, emp.Name);
            worksheet.SetCell(row, 3, emp.Department);
            worksheet.SetCell(row, 4, emp.Salary).AsCurrency();
            worksheet.SetCell(row, 5, emp.HireDate).AsDate();
            row++;
        }

        // Style the data range
        var dataRange = worksheet.GetRange(4, 1, row - 1, 5);
        dataRange.WithBorder();

        // Add totals
        worksheet.SetCell(row + 1, 3, "Total Salaries:")
            .Bold();
        worksheet.SetCell(row + 1, 4, employees.Sum(e => e.Salary))
            .AsCurrency()
            .WithTheme(FluentTheme.Total);

        worksheet.AutoFitColumns()
                .FreezeTopRow();

        workbook.SaveAs("employee-data.xlsx");
        Console.WriteLine("   ✓ Created employee-data.xlsx");
    }

    static void AdvancedStylingDemo()
    {
        Console.WriteLine("\n3. Advanced Styling Demo - Showcasing formatting features");

        using var workbook = FluentWorkbook.Create();
        
        var worksheet = workbook.AddWorksheet("Styling Demo");

        // Title with merged cells
        var titleRange = worksheet.GetRange("A1:E1");
        titleRange.WithValue("Advanced Styling Demonstration")
                  .Merge()
                  .Bold()
                  .WithFontSize(20)
                  .WithBackgroundColor(XLColor.DarkBlue)
                  .WithFontColor(XLColor.White)
                  .Center();

        // Different color themes
        worksheet.SetCell("A3", "Success").WithTheme(FluentTheme.Success);
        worksheet.SetCell("B3", "Warning").WithTheme(FluentTheme.Warning);
        worksheet.SetCell("C3", "Header").WithTheme(FluentTheme.Header);
        worksheet.SetCell("D3", "Data").WithTheme(FluentTheme.Data);
        worksheet.SetCell("E3", "Total").WithTheme(FluentTheme.Total);

        // Number formatting examples
        worksheet.SetCell("A5", "Currency:");
        worksheet.SetCell("B5", 1234.56).AsCurrency();

        worksheet.SetCell("A6", "Percentage:");
        worksheet.SetCell("B6", 0.75).AsPercentage();

        worksheet.SetCell("A7", "Date:");
        worksheet.SetCell("B7", DateTime.Now).AsDate();

        worksheet.SetCell("A8", "Custom Format:");
        worksheet.SetCell("B8", 42).WithNumberFormat("000000");

        // Border examples
        var borderDemo = worksheet.GetRange("A10:C12");
        borderDemo.WithValue("Border Demo")
                  .WithBorder()
                  .WithInsideBorder()
                  .Center();

        worksheet.AutoFitColumns();

        workbook.SaveAs("styling-demo.xlsx");
        Console.WriteLine("   ✓ Created styling-demo.xlsx");
    }

    static void FinancialReportDemo()
    {
        Console.WriteLine("\n4. Financial Report Demo - Creating a financial report");

        var financialData = new[]
        {
            new FinancialRecord { Account = "Revenue", Q1 = 100000, Q2 = 120000, Q3 = 135000, Q4 = 150000 },
            new FinancialRecord { Account = "Cost of Goods Sold", Q1 = -40000, Q2 = -48000, Q3 = -54000, Q4 = -60000 },
            new FinancialRecord { Account = "Operating Expenses", Q1 = -30000, Q2 = -32000, Q3 = -35000, Q4 = -38000 },
            new FinancialRecord { Account = "Marketing", Q1 = -10000, Q2 = -12000, Q3 = -13000, Q4 = -15000 },
            new FinancialRecord { Account = "Net Income", Q1 = 20000, Q2 = 28000, Q3 = 33000, Q4 = 37000 }
        };

        using var workbook = financialData.ToFinancialReport("Quarterly Financial Report 2024");
        
        // Add some additional formatting for the specific financial report
        var worksheet = workbook.GetWorksheet(1);
        
        // Highlight revenue row in green
        worksheet.GetRange("A5:E5").WithBackgroundColor(XLColor.LightGreen);
        
        // Highlight expense rows in light red
        for (int row = 6; row <= 8; row++)
        {
            worksheet.GetRange($"A{row}:E{row}").WithBackgroundColor(XLColor.LightPink);
        }
        
        // Highlight net income in dark green
        worksheet.GetRange("A9:E9").WithBackgroundColor(XLColor.DarkGreen).WithFontColor(XLColor.White);

        workbook.SaveAs("financial-report.xlsx");
        Console.WriteLine("   ✓ Created financial-report.xlsx");
    }

    static void ExtensionMethodsDemo()
    {
        Console.WriteLine("\n5. Extension Methods Demo - Using convenience extensions");

        var salesData = new[]
        {
            new SaleRecord { Product = "Laptop", Category = "Electronics", Price = 999.99m, Quantity = 5 },
            new SaleRecord { Product = "Mouse", Category = "Electronics", Price = 29.99m, Quantity = 20 },
            new SaleRecord { Product = "Keyboard", Category = "Electronics", Price = 79.99m, Quantity = 15 },
            new SaleRecord { Product = "Monitor", Category = "Electronics", Price = 299.99m, Quantity = 8 },
            new SaleRecord { Product = "Desk Chair", Category = "Furniture", Price = 199.99m, Quantity = 10 }
        };

        // Using extension method to quickly create a workbook from data
        using var workbook = salesData.ToFluentWorkbook("Sales Data", createTable: true);
        
        // Add a summary worksheet
        var summarySheet = workbook.AddWorksheet("Summary");
        
        summarySheet.SetCell("A1", "Sales Summary")
            .WithTheme(FluentTheme.Header)
            .WithFontSize(16);

        var totalRevenue = salesData.Sum(s => s.Price * s.Quantity);
        var totalQuantity = salesData.Sum(s => s.Quantity);

        summarySheet.SetCell("A3", "Total Revenue:");
        summarySheet.SetCell("B3", totalRevenue).AsCurrency().WithTheme(FluentTheme.Success);

        summarySheet.SetCell("A4", "Total Items Sold:");
        summarySheet.SetCell("B4", totalQuantity).WithTheme(FluentTheme.Success);

        summarySheet.SetCell("A5", "Average Order Value:");
        summarySheet.SetCell("B5", totalRevenue / salesData.Length).AsCurrency();

        summarySheet.AutoFitColumns();

        workbook.SaveAs("sales-with-extensions.xlsx");
        Console.WriteLine("   ✓ Created sales-with-extensions.xlsx");
    }

    static void FormulaDemo()
    {
        Console.WriteLine("\n6. Formula Demo - Showcasing formula capabilities");

        using var workbook = FluentWorkbook.Create();
        var worksheet = workbook.AddWorksheet("Formula Demo");

        // Title
        worksheet.SetCell("A1", "Excel Formula Demonstrations")
            .Bold()
            .WithFontSize(16)
            .WithFontColor(XLColor.DarkBlue);

        // Sample data for formulas
        worksheet.SetCell("A3", "Sample Data");
        worksheet.SetCell("A4", "Value 1:").Bold();
        worksheet.SetCell("B4", 100);
        worksheet.SetCell("A5", "Value 2:").Bold();
        worksheet.SetCell("B5", 250);
        worksheet.SetCell("A6", "Value 3:").Bold();
        worksheet.SetCell("B6", 75);

        // Basic arithmetic formulas
        worksheet.SetCell("A8", "Basic Arithmetic:").Bold().WithFontSize(12);
        
        worksheet.SetCell("A9", "Addition:");
        worksheet.SetCell("B9", "").WithAdd("B4", "B5").AsCurrency();
        
        worksheet.SetCell("A10", "Subtraction:");
        worksheet.SetCell("B10", "").WithSubtract("B5", "B6").AsCurrency();
        
        worksheet.SetCell("A11", "Multiplication:");
        worksheet.SetCell("B11", "").WithMultiply("B4", "B6").AsCurrency();
        
        worksheet.SetCell("A12", "Division:");
        worksheet.SetCell("B12", "").WithDivide("B5", "B4").WithNumberFormat("0.00");

        // Statistical formulas
        worksheet.SetCell("A14", "Statistical Functions:").Bold().WithFontSize(12);
        
        worksheet.SetCell("A15", "Sum:");
        worksheet.SetCell("B15", "").WithSum("B4:B6").AsCurrency();
        
        worksheet.SetCell("A16", "Average:");
        worksheet.SetCell("B16", "").WithAverage("B4:B6").WithNumberFormat("0.00");
        
        worksheet.SetCell("A17", "Maximum:");
        worksheet.SetCell("B17", "").WithMax("B4:B6").AsCurrency();
        
        worksheet.SetCell("A18", "Minimum:");
        worksheet.SetCell("B18", "").WithMin("B4:B6").AsCurrency();
        
        worksheet.SetCell("A19", "Count:");
        worksheet.SetCell("B19", "").WithCount("B4:B6");

        // Date and time functions
        worksheet.SetCell("A21", "Date & Time Functions:").Bold().WithFontSize(12);
        
        worksheet.SetCell("A22", "Today:");
        worksheet.SetCell("B22", "").WithToday().AsDate();
        
        worksheet.SetCell("A23", "Now:");
        worksheet.SetCell("B23", "").WithNow().WithNumberFormat("mm/dd/yyyy hh:mm");

        // Conditional formulas
        worksheet.SetCell("A25", "Conditional Functions:").Bold().WithFontSize(12);
        
        worksheet.SetCell("A26", "IF (B4>B6):");
        worksheet.SetCell("B26", "").WithIf("B4>B6", "\"Greater\"", "\"Lesser\"");
        
        worksheet.SetCell("A27", "SUMIF (>100):");
        worksheet.SetCell("B27", "").WithSumIf("B4:B6", ">100").AsCurrency();
        
        worksheet.SetCell("A28", "COUNTIF (>100):");
        worksheet.SetCell("B28", "").WithCountIf("B4:B6", ">100");

        // Text functions
        worksheet.SetCell("A30", "Text Functions:").Bold().WithFontSize(12);
        
        worksheet.SetCell("A31", "Concatenation:");
        worksheet.SetCell("B31", "").WithConcatenate("\"Value: \"", "B4");

        // Custom formula
        worksheet.SetCell("A33", "Custom Formula:").Bold().WithFontSize(12);
        worksheet.SetCell("A34", "Complex calculation:");
        worksheet.SetCell("B34", "").WithFormula("ROUND((B4+B5+B6)/3*1.1,2)").AsCurrency();

        // VLOOKUP example setup
        worksheet.SetCell("D3", "VLOOKUP Example:").Bold().WithFontSize(12);
        worksheet.SetCell("D4", "Product").Bold();
        worksheet.SetCell("E4", "Price").Bold();
        worksheet.SetCell("D5", "Apple");
        worksheet.SetCell("E5", 1.50);
        worksheet.SetCell("D6", "Banana");
        worksheet.SetCell("E6", 0.75);
        worksheet.SetCell("D7", "Orange");
        worksheet.SetCell("E7", 2.00);
        
        worksheet.SetCell("D9", "Lookup Apple:");
        worksheet.SetCell("E9", "").WithVLookup("\"Apple\"", "D5:E7", 2, true).AsCurrency();

        // Round function
        worksheet.SetCell("A36", "Rounding:").Bold().WithFontSize(12);
        worksheet.SetCell("A37", "Round to 0 decimals:");
        worksheet.SetCell("B37", "").WithRound("B16", 0);

        // Style the formulas differently
        var formulaRange = worksheet.GetRange("B9:B37");
        formulaRange.WithBackgroundColor(XLColor.LightYellow);

        worksheet.AutoFitColumns();

        workbook.SaveAs("formula-demo.xlsx");
        Console.WriteLine("   ✓ Created formula-demo.xlsx with comprehensive formula examples");
    }

    static void PredefinedNumberFormatDemo()
    {
        Console.WriteLine("\n7. Predefined Number Format Demo - Showcasing ClosedXML's built-in formats");

        using var workbook = FluentWorkbook.Create();
        var worksheet = workbook.AddWorksheet("Number Formats");

        // Title
        worksheet.SetCell("A1", "ClosedXML Predefined Number Format Demonstrations")
            .Bold()
            .WithFontSize(16)
            .WithFontColor(XLColor.DarkBlue);

        // Sample values for formatting
        var sampleNumber = 1234.567;
        var sampleDate = new DateTime(2024, 12, 15, 14, 30, 45);
        var samplePercentage = 0.1234;
        var sampleCurrency = 1234.56;
        var sampleFraction = 0.75;

        int row = 3;

        // Number Formats Section
        worksheet.SetCell($"A{row}", "NUMBER FORMATS").Bold().WithFontSize(14);
        worksheet.SetCell($"B{row}", "Sample Value").Bold();
        worksheet.SetCell($"C{row}", "Formatted Result").Bold();
        worksheet.SetCell($"D{row}", "Format Description").Bold();
        row++;

        // General format
        worksheet.SetCell($"A{row}", "General");
        worksheet.SetCell($"B{row}", sampleNumber);
        worksheet.SetCell($"C{row}", sampleNumber).AsGeneral();
        worksheet.SetCell($"D{row}", "Default Excel format");
        row++;

        // Number formats
        worksheet.SetCell($"A{row}", "Number (0)");
        worksheet.SetCell($"B{row}", sampleNumber);
        worksheet.SetCell($"C{row}", sampleNumber).AsNumber();
        worksheet.SetCell($"D{row}", "Integer format");
        row++;

        worksheet.SetCell($"A{row}", "Number (0.00)");
        worksheet.SetCell($"B{row}", sampleNumber);
        worksheet.SetCell($"C{row}", sampleNumber).AsNumberWith2Decimals();
        worksheet.SetCell($"D{row}", "Two decimal places");
        row++;

        worksheet.SetCell($"A{row}", "Number (#,##0)");
        worksheet.SetCell($"B{row}", sampleNumber);
        worksheet.SetCell($"C{row}", sampleNumber).AsNumberWithThousandsSeparator();
        worksheet.SetCell($"D{row}", "Thousands separator");
        row++;

        worksheet.SetCell($"A{row}", "Number (#,##0.00)");
        worksheet.SetCell($"B{row}", sampleNumber);
        worksheet.SetCell($"C{row}", sampleNumber).AsNumberWithThousandsSeparatorAnd2Decimals();
        worksheet.SetCell($"D{row}", "Thousands separator + 2 decimals");
        row++;

        // Custom decimal places
        worksheet.SetCell($"A{row}", "Number (4 decimals)");
        worksheet.SetCell($"B{row}", sampleNumber);
        worksheet.SetCell($"C{row}", sampleNumber).AsNumberWithDecimals(4);
        worksheet.SetCell($"D{row}", "Custom 4 decimal places");
        row++;

        row++; // Empty row

        // Currency Formats Section
        worksheet.SetCell($"A{row}", "CURRENCY FORMATS").Bold().WithFontSize(14);
        row++;

        worksheet.SetCell($"A{row}", "Currency (Built-in)");
        worksheet.SetCell($"B{row}", sampleCurrency);
        worksheet.SetCell($"C{row}", sampleCurrency).AsCurrencyBuiltIn();
        worksheet.SetCell($"D{row}", "Standard currency format");
        row++;

        worksheet.SetCell($"A{row}", "Currency (2 decimals)");
        worksheet.SetCell($"B{row}", sampleCurrency);
        worksheet.SetCell($"C{row}", sampleCurrency).AsCurrencyWith2Decimals();
        worksheet.SetCell($"D{row}", "Currency with 2 decimals");
        row++;

        worksheet.SetCell($"A{row}", "Currency (Red negatives)");
        worksheet.SetCell($"B{row}", -sampleCurrency);
        worksheet.SetCell($"C{row}", -sampleCurrency).AsCurrencyWithRedNegatives();
        worksheet.SetCell($"D{row}", "Negative values in red");
        row++;

        worksheet.SetCell($"A{row}", "Custom Currency (€)");
        worksheet.SetCell($"B{row}", sampleCurrency);
        worksheet.SetCell($"C{row}", sampleCurrency).AsCurrencyWithDecimals(2, "€");
        worksheet.SetCell($"D{row}", "Euro currency symbol");
        row++;

        worksheet.SetCell($"A{row}", "Accounting");
        worksheet.SetCell($"B{row}", sampleCurrency);
        worksheet.SetCell($"C{row}", sampleCurrency).AsAccounting();
        worksheet.SetCell($"D{row}", "Accounting format");
        row++;

        row++; // Empty row

        // Percentage Formats Section
        worksheet.SetCell($"A{row}", "PERCENTAGE FORMATS").Bold().WithFontSize(14);
        row++;

        worksheet.SetCell($"A{row}", "Percentage (0%)");
        worksheet.SetCell($"B{row}", samplePercentage);
        worksheet.SetCell($"C{row}", samplePercentage).AsPercentageBuiltIn();
        worksheet.SetCell($"D{row}", "Basic percentage");
        row++;

        worksheet.SetCell($"A{row}", "Percentage (0.00%)");
        worksheet.SetCell($"B{row}", samplePercentage);
        worksheet.SetCell($"C{row}", samplePercentage).AsPercentageWith2Decimals();
        worksheet.SetCell($"D{row}", "Percentage with 2 decimals");
        row++;

        worksheet.SetCell($"A{row}", "Percentage (4 decimals)");
        worksheet.SetCell($"B{row}", samplePercentage);
        worksheet.SetCell($"C{row}", samplePercentage).AsPercentageWithDecimals(4);
        worksheet.SetCell($"D{row}", "Percentage with 4 decimals");
        row++;

        row++; // Empty row

        // Date and Time Formats Section
        worksheet.SetCell($"A{row}", "DATE & TIME FORMATS").Bold().WithFontSize(14);
        row++;

        worksheet.SetCell($"A{row}", "Short Date");
        worksheet.SetCell($"B{row}", sampleDate);
        worksheet.SetCell($"C{row}", sampleDate).AsShortDate();
        worksheet.SetCell($"D{row}", "mm/dd/yyyy format");
        row++;

        worksheet.SetCell($"A{row}", "Date (d-mmm-yy)");
        worksheet.SetCell($"B{row}", sampleDate);
        worksheet.SetCell($"C{row}", sampleDate).AsDateDMmmYy();
        worksheet.SetCell($"D{row}", "Day-Month-Year format");
        row++;

        worksheet.SetCell($"A{row}", "Date (d-mmm)");
        worksheet.SetCell($"B{row}", sampleDate);
        worksheet.SetCell($"C{row}", sampleDate).AsDateDMmm();
        worksheet.SetCell($"D{row}", "Day-Month format");
        row++;

        worksheet.SetCell($"A{row}", "Date (mmm-yy)");
        worksheet.SetCell($"B{row}", sampleDate);
        worksheet.SetCell($"C{row}", sampleDate).AsDateMmmYy();
        worksheet.SetCell($"D{row}", "Month-Year format");
        row++;

        worksheet.SetCell($"A{row}", "Time (AM/PM)");
        worksheet.SetCell($"B{row}", sampleDate);
        worksheet.SetCell($"C{row}", sampleDate).AsTimeAmPm();
        worksheet.SetCell($"D{row}", "12-hour time format");
        row++;

        worksheet.SetCell($"A{row}", "Time (24-hour)");
        worksheet.SetCell($"B{row}", sampleDate);
        worksheet.SetCell($"C{row}", sampleDate).AsTime24Hour();
        worksheet.SetCell($"D{row}", "24-hour time format");
        row++;

        worksheet.SetCell($"A{row}", "DateTime");
        worksheet.SetCell($"B{row}", sampleDate);
        worksheet.SetCell($"C{row}", sampleDate).AsDateTime();
        worksheet.SetCell($"D{row}", "Date and time combined");
        row++;

        row++; // Empty row

        // Scientific and Special Formats Section
        worksheet.SetCell($"A{row}", "SPECIAL FORMATS").Bold().WithFontSize(14);
        row++;

        worksheet.SetCell($"A{row}", "Scientific");
        worksheet.SetCell($"B{row}", sampleNumber * 1000);
        worksheet.SetCell($"C{row}", sampleNumber * 1000).AsScientific();
        worksheet.SetCell($"D{row}", "Scientific notation");
        row++;

        worksheet.SetCell($"A{row}", "Fraction");
        worksheet.SetCell($"B{row}", sampleFraction);
        worksheet.SetCell($"C{row}", sampleFraction).AsFraction();
        worksheet.SetCell($"D{row}", "Fraction format");
        row++;

        worksheet.SetCell($"A{row}", "Text");
        worksheet.SetCell($"B{row}", sampleNumber);
        worksheet.SetCell($"C{row}", sampleNumber).AsText();
        worksheet.SetCell($"D{row}", "Treats numbers as text");
        row++;

        // Elapsed time example
        var elapsedTime = TimeSpan.FromHours(25.5); // 25 hours 30 minutes
        worksheet.SetCell($"A{row}", "Elapsed Time");
        worksheet.SetCell($"B{row}", elapsedTime.TotalDays); // Convert to Excel date value
        worksheet.SetCell($"C{row}", elapsedTime.TotalDays).AsElapsedTimeHours();
        worksheet.SetCell($"D{row}", "Elapsed time in hours");
        row++;

        row++; // Empty row

        // Range formatting example
        worksheet.SetCell($"A{row}", "RANGE FORMATTING EXAMPLE").Bold().WithFontSize(14);
        row++;

        worksheet.SetCell($"A{row}", "Sample Range:");
        var dataRange = worksheet.GetRange($"B{row}:E{row}");
        dataRange.WithValues(new object[,] { { 123.456, 0.789, 9876.54, new DateTime(2024, 6, 15) } });
        
        // Apply different formats to the range
        worksheet.GetCell($"B{row}").AsNumberWith2Decimals();
        worksheet.GetCell($"C{row}").AsPercentageWith2Decimals();
        worksheet.GetCell($"D{row}").AsCurrencyWith2Decimals();
        worksheet.GetCell($"E{row}").AsShortDate();
        
        row++;
        worksheet.SetCell($"A{row}", "Formats Applied:");
        worksheet.SetCell($"B{row}", "Number (2 dec)");
        worksheet.SetCell($"C{row}", "Percentage");
        worksheet.SetCell($"D{row}", "Currency");
        worksheet.SetCell($"E{row}", "Short Date");

        // Style the headers
        var headerRange = worksheet.GetRange("A3:D3");
        headerRange.WithTheme(FluentTheme.Header);

        // Auto-fit columns
        worksheet.AutoFitColumns();

        workbook.SaveAs("predefined-number-formats-demo.xlsx");
        Console.WriteLine("   ✓ Created predefined-number-formats-demo.xlsx showcasing ClosedXML's built-in formats");
    }

    static void RangeAddressDemo()
    {
        Console.WriteLine("\n8. Range Address Demo - Getting range addresses as strings");

        using var workbook = FluentWorkbook.Create();
        var worksheet = workbook.AddWorksheet("Range Addresses");

        // Title
        worksheet.SetCell("A1", "Range Address Demonstration")
            .Bold()
            .WithFontSize(16)
            .WithFontColor(XLColor.DarkBlue);

        // Example 1: Get address of a range created with coordinates
        var range1 = worksheet.GetRange(3, 1, 5, 3); // From A3 to C5
        var address1 = range1.GetAddress();
        
        worksheet.SetCell("A3", "Range created with coordinates (3,1,5,3):");
        worksheet.SetCell("A4", $"Address: {address1}").WithFontColor(XLColor.DarkGreen);

        // Example 2: Get address of a range created with string notation
        var range2 = worksheet.GetRange("E3:G7");
        var address2 = range2.GetAddress();
        
        worksheet.SetCell("A6", "Range created with string notation (E3:G7):");
        worksheet.SetCell("A7", $"Address: {address2}").WithFontColor(XLColor.DarkGreen);

        // Example 3: Practical usage - highlight ranges and show their addresses
        worksheet.SetCell("A9", "Practical Example - Highlighted Ranges:").Bold();

        // Create some sample ranges and highlight them
        var salesRange = worksheet.GetRange(11, 1, 15, 4);
        salesRange.WithValue("Sales Data")
                  .WithBackgroundColor(XLColor.LightBlue)
                  .Center();
        
        var totalsRange = worksheet.GetRange(17, 1, 17, 4);
        totalsRange.WithValue("Totals")
                   .WithBackgroundColor(XLColor.LightGreen)
                   .Bold()
                   .Center();

        // Display the addresses
        worksheet.SetCell("A20", "Sales Data Range Address:")
                 .Bold();
        worksheet.SetCell("B20", salesRange.GetAddress())
                 .WithFontColor(XLColor.Blue);

        worksheet.SetCell("A21", "Totals Range Address:")
                 .Bold();
        worksheet.SetCell("B21", totalsRange.GetAddress())
                 .WithFontColor(XLColor.Green);

        // Example 4: Dynamic range creation and address retrieval
        worksheet.SetCell("A23", "Dynamic Range Example:").Bold();
        
        // Create a range dynamically based on data
        var dataStartRow = 25;
        var dataEndRow = dataStartRow + 3;
        var dataStartCol = 1;
        var dataEndCol = 3;
        
        var dynamicRange = worksheet.GetRange(dataStartRow, dataStartCol, dataEndRow, dataEndCol);
        dynamicRange.WithValue("Dynamic Data")
                    .WithBackgroundColor(XLColor.LightYellow)
                    .WithBorder();

        worksheet.SetCell("A30", "Dynamic Range Address:")
                 .Bold();
        worksheet.SetCell("B30", dynamicRange.GetAddress())
                 .WithFontColor(XLColor.Orange);

        // Show how this can be useful for debugging or logging
        worksheet.SetCell("A32", "Use Cases:").Bold().WithFontSize(12);
        worksheet.SetCell("A33", "• Debugging - Know exactly which range you're working with");
        worksheet.SetCell("A34", "• Logging - Record which ranges were processed");
        worksheet.SetCell("A35", "• Dynamic formulas - Build formulas using range addresses");
        worksheet.SetCell("A36", "• Range validation - Verify range coordinates are correct");

        worksheet.AutoFitColumns();

        workbook.SaveAs("range-address-demo.xlsx");
        Console.WriteLine("   ✓ Created range-address-demo.xlsx demonstrating GetAddress() method");
    }

    static void RangeFormulaAndMergeDemo()
    {
        Console.WriteLine("\n9. Range Formula and Merge Demo - Merging cells and setting formulas");

        using var workbook = FluentWorkbook.Create();
        var worksheet = workbook.AddWorksheet("Range Formula & Merge");

        // Title
        worksheet.SetCell("A1", "Range Formula and Merge Demonstration")
            .Bold()
            .WithFontSize(16)
            .WithFontColor(XLColor.DarkBlue);

        // Example 1: Your desired syntax - merge and set formula
        worksheet.SetCell("A3", "Example 1: Your exact syntax");
        
        // IMPORTANT: Note the coordinate correction - Excel uses 1-based indexing
        // Your example used (0,0,0,1) which would be invalid
        // Corrected to (4,1,4,2) which means row 4, columns A to B
        var mergedRange1 = worksheet.GetRange(4, 1, 4, 2); // A4:B4
        mergedRange1.Merge().WithFormula("1+2");
        
        worksheet.SetCell("A5", "Range A4:B4 merged with formula '1+2'").WithFontColor(XLColor.DarkGreen);

        // Example 2: Alternative syntax - set value first, then formula
        worksheet.SetCell("A7", "Example 2: Alternative syntax");
        
        var mergedRange2 = worksheet.GetRange(8, 1, 8, 3); // A8:C8
        mergedRange2.WithValue(0).WithFormula("10*5").Merge();
        
        worksheet.SetCell("A9", "Range A8:C8 with value first, then formula '10*5', then merged").WithFontColor(XLColor.DarkGreen);

        // Example 3: More complex formula
        worksheet.SetCell("A11", "Example 3: Complex formula with merge");
        
        var mergedRange3 = worksheet.GetRange(12, 1, 12, 4); // A12:D12
        mergedRange3.Merge().WithFormula("SUM(1,2,3,4,5)").Bold().Center();
        
        worksheet.SetCell("A13", "Range A12:D12 merged with SUM formula and styling").WithFontColor(XLColor.DarkGreen);

        // Example 4: Using string addresses instead of coordinates
        worksheet.SetCell("A15", "Example 4: Using string addresses");
        
        var mergedRange4 = worksheet.GetRange("A16:C16");
        mergedRange4.Merge()
                    .WithFormula("AVERAGE(10,20,30)")
                    .WithBackgroundColor(XLColor.LightYellow)
                    .Center();
        
        worksheet.SetCell("A17", "Range A16:C16 merged with AVERAGE formula and background").WithFontColor(XLColor.DarkGreen);

        // Example 5: Formula with cell references
        worksheet.SetCell("A19", "Example 5: Formula with cell references");
        
        // First, set some values to reference
        worksheet.SetCell("E20", 100);
        worksheet.SetCell("F20", 200);
        worksheet.SetCell("G20", 300);
        
        var mergedRange5 = worksheet.GetRange(21, 1, 21, 3); // A21:C21
        mergedRange5.Merge().WithFormula("SUM(E20:G20)").AsCurrency();
        
        worksheet.SetCell("A22", "Range A21:C21 merged with formula referencing E20:G20").WithFontColor(XLColor.DarkGreen);

        // Example 6: Conditional formula
        worksheet.SetCell("A24", "Example 6: Conditional IF formula");
        
        var mergedRange6 = worksheet.GetRange(25, 1, 25, 2); // A25:B25
        mergedRange6.Merge().WithFormula("IF(E20>150,\"High\",\"Low\")").Italic();
        
        worksheet.SetCell("A26", "Range A25:B25 merged with IF formula").WithFontColor(XLColor.DarkGreen);

        // Example 7: Built-in formatting methods for merged ranges
        worksheet.SetCell("A28", "BUILT-IN FORMATTING EXAMPLES").Bold().WithFontSize(14);
        var formattingHeaderRange = worksheet.GetRange("A28:C28");
        formattingHeaderRange.WithBackgroundColor(XLColor.LightBlue);
        
        // Number with thousands separator
        var mergedRange7a = worksheet.GetRange(29, 1, 29, 2); // A29:B29
        mergedRange7a.Merge().WithFormula("1000+2000+3000").AsNumberWithThousandsSeparator().Center();
        worksheet.SetCell("C29", "AsNumberWithThousandsSeparator()").WithFontColor(XLColor.DarkGreen);
        
        // Currency with 2 decimals
        var mergedRange7b = worksheet.GetRange(30, 1, 30, 2); // A30:B30
        mergedRange7b.Merge().WithFormula("1234.567").AsCurrencyWith2Decimals().Center();
        worksheet.SetCell("C30", "AsCurrencyWith2Decimals()").WithFontColor(XLColor.DarkGreen);
        
        // Percentage with decimals
        var mergedRange7c = worksheet.GetRange(31, 1, 31, 2); // A31:B31
        mergedRange7c.Merge().WithFormula("0.12345").AsPercentageWithDecimals(3).Center();
        worksheet.SetCell("C31", "AsPercentageWithDecimals(3)").WithFontColor(XLColor.DarkGreen);
        
        // Scientific notation
        var mergedRange7d = worksheet.GetRange(32, 1, 32, 2); // A32:B32
        mergedRange7d.Merge().WithFormula("123456789").AsScientific().Center();
        worksheet.SetCell("C32", "AsScientific()").WithFontColor(XLColor.DarkGreen);
        
        // Accounting format
        var mergedRange7e = worksheet.GetRange(33, 1, 33, 2); // A33:B33
        mergedRange7e.Merge().WithFormula("9876.54").AsAccounting().Center();
        worksheet.SetCell("C33", "AsAccounting()").WithFontColor(XLColor.DarkGreen);
        
        // Date format
        var mergedRange7f = worksheet.GetRange(34, 1, 34, 2); // A34:B34
        mergedRange7f.Merge().WithFormula("TODAY()").AsShortDate().Center();
        worksheet.SetCell("C34", "AsShortDate()").WithFontColor(XLColor.DarkGreen);

        // Example 8: Show the coordinate correction for your original example
        worksheet.SetCell("A36", "Coordinate Information:").Bold().WithFontSize(12);
        worksheet.SetCell("A37", "• Excel uses 1-based indexing (row 1, column 1 = A1)");
        worksheet.SetCell("A38", "• Your example (0,0,0,1) is invalid in 1-based system");
        worksheet.SetCell("A39", "• Corrected: (1,1,1,2) = A1:B1 or (4,1,4,2) = A4:B4");
        worksheet.SetCell("A40", "• For a 2-cell horizontal merge: GetRange(row, col, row, col+1)");
        worksheet.SetCell("A41", "• All built-in formatting methods from FluentCell are available!");

        // Style the information section
        var infoRange = worksheet.GetRange("A36:A41");
        infoRange.WithBackgroundColor(XLColor.LightGray);

        // Add headers for the values we used
        worksheet.SetCell("E19", "Sample Values:").Bold();
        worksheet.SetCell("E20", 100).Bold();
        worksheet.SetCell("F20", 200).Bold();
        worksheet.SetCell("G20", 300).Bold();

        worksheet.AutoFitColumns();

        workbook.SaveAs("range-formula-merge-demo.xlsx");
        Console.WriteLine("   ✓ Created range-formula-merge-demo.xlsx demonstrating range formulas and merging");
    }

    static void HorizontalRangeValuesDemo()
    {
        Console.WriteLine("\n10. Horizontal Range Values Demo - Testing automatic orientation detection");

        using var workbook = FluentWorkbook.Create();
        var worksheet = workbook.AddWorksheet("Horizontal Range Values");

        // Title
        worksheet.SetCell("A1", "Horizontal Range Values Demonstration")
            .Bold()
            .WithFontSize(16)
            .WithFontColor(XLColor.DarkBlue);

        // Demonstrate the issue and fix
        worksheet.SetCell("A3", "THE ISSUE (before fix):").Bold().WithFontSize(12);
        worksheet.SetCell("A4", "• Creating a horizontal range like B5:E5");
        worksheet.SetCell("A5", "• Using WithValues(array) would place values in first column only");
        worksheet.SetCell("A6", "• Values would appear in B5, A6, A7, A8 instead of B5, C5, D5, E5");

        worksheet.SetCell("A8", "THE FIX (automatic orientation detection):").Bold().WithFontSize(12);
        worksheet.SetCell("A9", "• Range dimensions are analyzed");
        worksheet.SetCell("A10", "• Single row + multiple columns = horizontal orientation");
        worksheet.SetCell("A11", "• Single column + multiple rows = vertical orientation");
        worksheet.SetCell("A12", "• Square/rectangular ranges default to vertical");

        // Example 1: Horizontal range (user's scenario)
        worksheet.SetCell("A14", "Example 1: Horizontal Range (B15:E15)").Bold().WithFontColor(XLColor.DarkGreen);
        var horizontalRange = worksheet.GetRange("B15:E15");
        var horizontalData = new object[] { "Q1", "Q2", "Q3", "Q4" };
        horizontalRange.WithValues(horizontalData)
                      .WithBackgroundColor(XLColor.LightBlue)
                      .Bold()
                      .Center();

        worksheet.SetCell("A16", "✓ Values correctly placed horizontally").WithFontColor(XLColor.DarkGreen);
        worksheet.SetCell("A17", $"Range address: {horizontalRange.GetAddress()}").WithFontColor(XLColor.Blue);

        // Example 2: Another horizontal range in a different location
        worksheet.SetCell("A19", "Example 2: Horizontal Range (C20:F20)").Bold().WithFontColor(XLColor.DarkGreen);
        var horizontalRange2 = worksheet.GetRange("C20:F20");
        var salesData = new object[] { 100, 200, 150, 300 };
        horizontalRange2.WithValues(salesData)
                       .AsCurrency()
                       .WithBackgroundColor(XLColor.LightGreen)
                       .Center();

        worksheet.SetCell("A21", "✓ Sales values placed horizontally with currency formatting").WithFontColor(XLColor.DarkGreen);

        // Example 3: Vertical range (should work as before)
        worksheet.SetCell("A23", "Example 3: Vertical Range (A25:A28)").Bold().WithFontColor(XLColor.Purple);
        var verticalRange = worksheet.GetRange("A25:A28");
        var verticalData = new object[] { "Product A", "Product B", "Product C", "Product D" };
        verticalRange.WithValues(verticalData)
                     .WithBackgroundColor(XLColor.LightYellow)
                     .Bold();

        worksheet.SetCell("B25", "✓ Vertical placement still works").WithFontColor(XLColor.Purple);

        // Example 4: Explicit orientation control
        worksheet.SetCell("A30", "Example 4: Explicit Orientation Control").Bold().WithFontSize(12);
        
        worksheet.SetCell("A31", "Force Horizontal in Vertical Range (H32:H35):").Bold();
        var forceHorizontalRange = worksheet.GetRange("H32:H35");
        var monthData = new object[] { "Jan", "Feb", "Mar", "Apr" };
        forceHorizontalRange.WithValues(monthData, RangeOrientation.Horizontal)
                           .WithBackgroundColor(XLColor.LightCyan)
                           .Center();
        worksheet.SetCell("A33", "✓ Forced horizontal - only H32 filled, others ignored").WithFontColor(XLColor.Orange);

        worksheet.SetCell("A35", "Force Vertical in Horizontal Range (B36:E36):").Bold();
        var forceVerticalRange = worksheet.GetRange("B36:E36");
        var numberData = new object[] { 10, 20, 30, 40 };
        forceVerticalRange.WithValues(numberData, RangeOrientation.Vertical)
                         .AsNumber()
                         .WithBackgroundColor(XLColor.LightPink)
                         .Center();
        worksheet.SetCell("A37", "✓ Forced vertical - only B36 filled, others ignored").WithFontColor(XLColor.Orange);

        // Example 5: Real-world scenario
        worksheet.SetCell("A39", "Example 5: Real-World Header Scenario").Bold().WithFontSize(12);
        
        // Create headers for a table
        var headerRange = worksheet.GetRange("A41:D41");
        var headers = new object[] { "Employee", "Department", "Salary", "Rating" };
        headerRange.WithValues(headers)
                   .Bold()
                   .WithBackgroundColor(XLColor.DarkBlue)
                   .WithFontColor(XLColor.White)
                   .Center()
                   .WithBorder();

        // Add some sample data rows
        var employees = new[]
        {
            new object[] { "John Doe", "IT", 75000, 4.5 },
            new object[] { "Jane Smith", "HR", 65000, 4.8 },
            new object[] { "Bob Wilson", "Finance", 80000, 4.2 }
        };

        for (int i = 0; i < employees.Length; i++)
        {
            var rowRange = worksheet.GetRange(42 + i, 1, 42 + i, 4);
            rowRange.WithValues(employees[i])
                   .WithBorder();
            
            // Format salary column
            worksheet.GetCell(42 + i, 3).AsCurrency();
        }

        worksheet.SetCell("A46", "✓ Complete table with headers and data using horizontal ranges").WithFontColor(XLColor.DarkGreen);

        // Summary information
        worksheet.SetCell("A48", "SUMMARY:").Bold().WithFontSize(14);
        worksheet.SetCell("A49", "• Automatic orientation detection based on range dimensions");
        worksheet.SetCell("A50", "• Horizontal ranges (1 row × N columns) → horizontal placement");
        worksheet.SetCell("A51", "• Vertical ranges (N rows × 1 column) → vertical placement");
        worksheet.SetCell("A52", "• Explicit control with WithValues(data, RangeOrientation.Horizontal/Vertical)");
        worksheet.SetCell("A53", "• Backwards compatible - existing vertical ranges still work");

        // Highlight the user's specific scenario
        worksheet.SetCell("F48", "YOUR SCENARIO SOLVED:").Bold().WithFontSize(12).WithFontColor(XLColor.Red);
        worksheet.SetCell("F49", "var range = worksheet.GetRange(\"B5:E5\");");
        worksheet.SetCell("F50", "range.WithValues(new[] { \"Q1\", \"Q2\", \"Q3\", \"Q4\" });");
        worksheet.SetCell("F51", "// Now correctly places values horizontally!");
        
        var scenarioRange = worksheet.GetRange("F49:F51");
        scenarioRange.WithBackgroundColor(XLColor.Yellow);

        // Set column widths for the range demo
        worksheet.Worksheet.Column("A").Width = 12;
        worksheet.Worksheet.Column("C").Width = 20;

        worksheet.AutoFitColumns();

        workbook.SaveAs("horizontal-range-values-demo.xlsx");
        Console.WriteLine("   ✓ Created horizontal-range-values-demo.xlsx demonstrating the fix for horizontal range values");
    }

    static void BorderDemo()
    {
        Console.WriteLine("\n11. Border Demo - Showcasing new border capabilities");

        using var workbook = FluentWorkbook.Create();
        var worksheet = workbook.AddWorksheet("Border Demo");

        // Title
        worksheet.SetCell("A1", "Border Demonstration - Thick Outside, Thin Inside")
            .Bold()
            .WithFontSize(16)
            .WithFontColor(XLColor.DarkBlue);

        // Example 1: Your requested functionality - thick border around, thin inside
        worksheet.SetCell("A3", "Example 1: Table Borders (Thick Outside, Thin Inside)")
            .Bold()
            .WithFontSize(14)
            .WithFontColor(XLColor.DarkGreen);

        var tableRange = worksheet.GetRange("B5:F10");
        tableRange.WithValues(new object[,]
        {
            { "Product", "Q1", "Q2", "Q3", "Q4" },
            { "Laptops", 100, 120, 130, 150 },
            { "Mice", 200, 180, 220, 240 },
            { "Keyboards", 80, 90, 85, 95 },
            { "Monitors", 50, 60, 55, 70 },
            { "Total", 430, 450, 490, 555 }
        });

        // Apply the new table borders functionality
        tableRange.WithTableBorders() // This is your main feature!
                  .Center();

        // Format the header row
        var headerRange = worksheet.GetRange("B5:F5");
        headerRange.Bold()
                   .WithBackgroundColor(XLColor.LightBlue)
                   .WithFontColor(XLColor.DarkBlue);

        // Format the total row
        var totalRange = worksheet.GetRange("B10:F10");
        totalRange.Bold()
                  .WithBackgroundColor(XLColor.LightGreen);

        worksheet.SetCell("A11", "✓ WithTableBorders() - Thick outside, thin inside").WithFontColor(XLColor.DarkGreen);

        // Example 2: Custom border styles
        worksheet.SetCell("A13", "Example 2: Custom Border Styles")
            .Bold()
            .WithFontSize(14)
            .WithFontColor(XLColor.DarkGreen);

        var customRange = worksheet.GetRange("B15:E18");
        customRange.WithValues(new object[,]
        {
            { "Custom", "Border", "Style", "Demo" },
            { "Medium", "Outside", "Dashed", "Inside" },
            { "Different", "Colors", "And", "Styles" },
            { "Very", "Flexible", "Border", "Control" }
        });

        customRange.WithTableBorders(XLBorderStyleValues.Medium, XLBorderStyleValues.Dashed)
                   .WithBorderColor(XLColor.DarkRed)
                   .Center();

        worksheet.SetCell("A19", "✓ WithTableBorders(Medium, Dashed) + Red color").WithFontColor(XLColor.DarkGreen);

        // Example 3: Individual border style methods
        worksheet.SetCell("A21", "Example 3: Individual Border Style Methods")
            .Bold()
            .WithFontSize(14)
            .WithFontColor(XLColor.DarkGreen);

        // Thick border example
        var thickRange = worksheet.GetRange("B23:C24");
        thickRange.WithValue("Thick Border")
                  .WithThickBorder()
                  .WithBackgroundColor(XLColor.LightYellow)
                  .Center()
                  .Bold();

        // Double border example
        var doubleRange = worksheet.GetRange("E23:F24");
        doubleRange.WithValue("Double Border")
                   .WithDoubleBorder()
                   .WithBackgroundColor(XLColor.LightCyan)
                   .Center()
                   .Bold();

        // Dashed border example
        var dashedRange = worksheet.GetRange("B26:C27");
        dashedRange.WithValue("Dashed Border")
                   .WithDashedBorder()
                   .WithBackgroundColor(XLColor.LightPink)
                   .Center()
                   .Bold();

        // Dotted border example
        var dottedRange = worksheet.GetRange("E26:F27");
        dottedRange.WithValue("Dotted Border")
                   .WithDottedBorder()
                   .WithBackgroundColor(XLColor.LightGray)
                   .Center()
                   .Bold();

        worksheet.SetCell("A28", "✓ WithThickBorder(), WithDoubleBorder(), WithDashedBorder(), WithDottedBorder()").WithFontColor(XLColor.DarkGreen);

        // Example 4: Individual border sides
        worksheet.SetCell("A30", "Example 4: Individual Border Sides")
            .Bold()
            .WithFontSize(14)
            .WithFontColor(XLColor.DarkGreen);

        var individualRange = worksheet.GetRange("B32:E35");
        individualRange.WithValues(new object[,]
        {
            { "Top", "Only", "Thick", "Border" },
            { "Right", "Only", "Medium", "Border" },
            { "Bottom", "Only", "Double", "Border" },
            { "Left", "Only", "Dashed", "Border" }
        })
        .Center();

        // Apply individual borders to demonstrate the WithBorders method
        worksheet.GetRange("B32:E32").WithBorders(top: XLBorderStyleValues.Thick);
        worksheet.GetRange("E32:E35").WithBorders(right: XLBorderStyleValues.Medium);
        worksheet.GetRange("B35:E35").WithBorders(bottom: XLBorderStyleValues.Double);
        worksheet.GetRange("B32:B35").WithBorders(left: XLBorderStyleValues.Dashed);

        worksheet.SetCell("A36", "✓ WithBorders(top: Thick, right: Medium, etc.)").WithFontColor(XLColor.DarkGreen);

        // Example 5: No borders
        worksheet.SetCell("A38", "Example 5: Remove Borders")
            .Bold()
            .WithFontSize(14)
            .WithFontColor(XLColor.DarkGreen);

        var noBorderRange = worksheet.GetRange("B40:D42");
        noBorderRange.WithValue("No Borders")
                     .WithoutBorder()
                     .WithBackgroundColor(XLColor.LightGreen)
                     .Center()
                     .Bold();

        worksheet.SetCell("A43", "✓ WithoutBorder() - removes all borders").WithFontColor(XLColor.DarkGreen);

        // Summary of new methods
        worksheet.SetCell("A45", "NEW BORDER METHODS SUMMARY:")
            .Bold()
            .WithFontSize(16)
            .WithFontColor(XLColor.DarkBlue);

        var summaryData = new string[,]
        {
            { "Method", "Description", "Use Case" },
            { "WithTableBorders()", "Thick outside, thin inside", "Data tables, your main request!" },
            { "WithTableBorders(out, in)", "Custom outside/inside styles", "Flexible table formatting" },
            { "WithThickBorder()", "Thick border around range", "Emphasis, headers" },
            { "WithMediumBorder()", "Medium border around range", "Moderate emphasis" },
            { "WithDoubleBorder()", "Double line border", "Special formatting" },
            { "WithDashedBorder()", "Dashed border around range", "Temporary/draft content" },
            { "WithDottedBorder()", "Dotted border around range", "Subtle separation" },
            { "WithBorders(t,r,b,l)", "Individual side control", "Custom border patterns" },
            { "WithBorderColor(color)", "Set all border colors", "Color coordination" },
            { "WithoutBorder()", "Remove all borders", "Clean up formatting" }
        };

        var summaryRange = worksheet.GetRange("A47:C57");
        summaryRange.WithValues(summaryData)
                    .WithTableBorders()
                    .WithBorderColor(XLColor.DarkBlue);

        // Format summary header
        var summaryHeaderRange = worksheet.GetRange("A47:C47");
        summaryHeaderRange.Bold()
                          .WithBackgroundColor(XLColor.DarkBlue)
                          .WithFontColor(XLColor.White);

        // Your specific use case highlighted
        var yourUseCaseRange = worksheet.GetRange("A48:C48");
        yourUseCaseRange.WithBackgroundColor(XLColor.Yellow)
                        .Bold();

        worksheet.AutoFitColumns();

        workbook.SaveAs("border-demo.xlsx");
        Console.WriteLine("   ✓ Created border-demo.xlsx showcasing thick outside, thin inside borders and more!");
    }

    static void TextWrapDemo()
    {
        Console.WriteLine("\n12. Text Wrap Demo - Showcasing new text wrapping capabilities");

        using var workbook = FluentWorkbook.Create();
        var worksheet = workbook.AddWorksheet("Text Wrap Demo");

        // Title
        worksheet.SetCell("A1", "Text Wrapping Demonstration")
            .Bold()
            .WithFontSize(16)
            .WithFontColor(XLColor.DarkBlue);

        // Example 1: Basic text wrapping for cells
        worksheet.SetCell("A3", "Example 1: Cell Text Wrapping")
            .Bold()
            .WithFontSize(14)
            .WithFontColor(XLColor.DarkGreen);

        // Set column width to demonstrate wrapping
        worksheet.Worksheet.Column("B").Width = 15;
        worksheet.Worksheet.Column("D").Width = 15;

        // Cell with text wrap enabled
        var longText = "This is a very long text that should wrap to multiple lines when text wrapping is enabled. It demonstrates how the WithTextWrap() method works.";
        
        worksheet.SetCell("A5", "With Text Wrap:").Bold();
        worksheet.SetCell("B5", longText)
            .WithTextWrap()
            .WithBorder()
            .WithBackgroundColor(XLColor.LightYellow);

        // Cell without text wrap for comparison
        worksheet.SetCell("A7", "Without Text Wrap:").Bold();
        worksheet.SetCell("B7", longText)
            .WithoutTextWrap()
            .WithBorder()
            .WithBackgroundColor(XLColor.LightPink);

        worksheet.SetCell("A9", "✓ WithTextWrap() vs WithoutTextWrap()").WithFontColor(XLColor.DarkGreen);

        // Example 2: Range text wrapping
        worksheet.SetCell("A11", "Example 2: Range Text Wrapping")
            .Bold()
            .WithFontSize(14)
            .WithFontColor(XLColor.DarkGreen);

        var textData = new object[,]
        {
            { "Header 1", "Header 2", "Header 3" },
            { "Short text", "This is a medium length text that should wrap nicely", "Another example of wrapping text" },
            { "Brief", "Long description that explains the product features in detail and why customers should choose it", "Conclusion and summary" }
        };

        var rangeWithWrap = worksheet.GetRange("A13:C15");
        rangeWithWrap.WithValues(textData)
                     .WithTextWrap()
                     .WithBorder()
                     .WithInsideBorder()
                     .WithVerticalAlignment(XLAlignmentVerticalValues.Top);

        // Format the header row
        var headerRange = worksheet.GetRange("A13:C13");
        headerRange.Bold()
                   .WithBackgroundColor(XLColor.LightBlue)
                   .WithFontColor(XLColor.DarkBlue)
                   .Center();

        // Set column widths for the range demo
        worksheet.Worksheet.Column("A").Width = 12;
        worksheet.Worksheet.Column("C").Width = 20;

        worksheet.SetCell("A16", "✓ Entire range with text wrapping enabled").WithFontColor(XLColor.DarkGreen);

        // Example 3: Conditional text wrapping
        worksheet.SetCell("A18", "Example 3: Conditional Text Wrapping")
            .Bold()
            .WithFontSize(14)
            .WithFontColor(XLColor.DarkGreen);

        worksheet.SetCell("A20", "Enable wrap (true):").Bold();
        worksheet.SetCell("B20", "This text will wrap because we pass true to WithTextWrap()")
            .WithTextWrap(true)
            .WithBorder()
            .WithBackgroundColor(XLColor.LightGreen);

        worksheet.SetCell("A22", "Disable wrap (false):").Bold();
        worksheet.SetCell("B22", "This text will not wrap because we pass false to WithTextWrap()")
            .WithTextWrap(false)
            .WithBorder()
            .WithBackgroundColor(XLColor.LightCoral);

        worksheet.SetCell("A24", "✓ WithTextWrap(bool) for conditional control").WithFontColor(XLColor.DarkGreen);

        // Example 4: Text wrap with other formatting
        worksheet.SetCell("A26", "Example 4: Text Wrap with Other Formatting")
            .Bold()
            .WithFontSize(14)
            .WithFontColor(XLColor.DarkGreen);

        var formattedRange = worksheet.GetRange("A28:B30");
        formattedRange.WithValues(new object[,]
        {
            { "Formatted Text Wrapping", "This demonstrates text wrapping combined with other formatting options" },
            { "Bold + Wrap", "Bold text that wraps nicely with background color and borders" },
            { "Currency + Wrap", "Financial data explanation: $1,234.56 represents the quarterly profit margin calculation" }
        });

        formattedRange.WithTextWrap()
                      .WithBorder()
                      .WithInsideBorder();

        // Apply different formatting to each row
        worksheet.GetRange("A28:B28")
            .Bold()
            .WithBackgroundColor(XLColor.DarkBlue)
            .WithFontColor(XLColor.White);

        worksheet.GetRange("A29:B29")
            .Bold()
            .WithBackgroundColor(XLColor.LightYellow);

        worksheet.GetRange("A30:B30")
            .WithBackgroundColor(XLColor.LightCyan);

        // Format the currency in the last cell
        worksheet.SetCell("B30", "Financial data explanation: $1,234.56 represents the quarterly profit margin calculation")
            .WithTextWrap();

        worksheet.SetCell("A31", "✓ Text wrap combined with colors, borders, and other formatting").WithFontColor(XLColor.DarkGreen);

        // Example 5: Real-world scenario - Product descriptions
        worksheet.SetCell("A33", "Example 5: Real-World Scenario - Product Catalog")
            .Bold()
            .WithFontSize(14)
            .WithFontColor(XLColor.DarkGreen);

        var productData = new object[,]
        {
            { "Product", "Price", "Description" },
            { "Laptop Pro X1", 1299.99, "High-performance laptop with Intel i7 processor, 16GB RAM, 512GB SSD, and 15.6-inch 4K display. Perfect for professionals and gamers." },
            { "Wireless Mouse", 49.99, "Ergonomic wireless mouse with precision tracking, 12-month battery life, and comfortable grip design for extended use." },
            { "Mechanical Keyboard", 129.99, "Premium mechanical keyboard with RGB backlighting, tactile switches, and programmable keys for enhanced productivity and gaming experience." }
        };

        var productRange = worksheet.GetRange("A35:C38");
        productRange.WithValues(productData)
                    .WithTextWrap()
                    .WithTableBorders()
                    .WithVerticalAlignment(XLAlignmentVerticalValues.Top);

        // Format header
        var productHeaderRange = worksheet.GetRange("A35:C35");
        productHeaderRange.Bold()
                          .WithBackgroundColor(XLColor.DarkGreen)
                          .WithFontColor(XLColor.White)
                          .Center();

        // Format price column as currency
        worksheet.GetRange("B36:B38").AsCurrency();

        // Set appropriate column widths
        worksheet.Worksheet.Column("A").Width = 18;
        worksheet.Worksheet.Column("B").Width = 12;
        worksheet.Worksheet.Column("C").Width = 50;

        worksheet.SetCell("A39", "✓ Product catalog with wrapped descriptions").WithFontColor(XLColor.DarkGreen);

        // Summary of new methods
        worksheet.SetCell("A41", "NEW TEXT WRAP METHODS SUMMARY:")
            .Bold()
            .WithFontSize(16)
            .WithFontColor(XLColor.DarkBlue);

        var summaryData = new string[,]
        {
            { "Method", "Description", "Use Case" },
            { "WithTextWrap()", "Enables text wrapping", "Long content that should wrap to multiple lines" },
            { "WithoutTextWrap()", "Disables text wrapping", "Ensure text stays on single line" },
            { "WithTextWrap(bool)", "Conditional text wrapping", "Enable/disable based on conditions or user preferences" }
        };

        var summaryRange = worksheet.GetRange("A43:C46");
        summaryRange.WithValues(summaryData)
                    .WithTableBorders()
                    .WithTextWrap()
                    .WithVerticalAlignment(XLAlignmentVerticalValues.Top);

        // Format summary header
        var summaryHeaderRange = worksheet.GetRange("A43:C43");
        summaryHeaderRange.Bold()
                          .WithBackgroundColor(XLColor.DarkBlue)
                          .WithFontColor(XLColor.White);

        // Additional tips
        worksheet.SetCell("A48", "TIPS FOR EFFECTIVE TEXT WRAPPING:")
            .Bold()
            .WithFontSize(12)
            .WithFontColor(XLColor.Purple);

        var tips = new[]
        {
            "• Set appropriate column widths for best wrapping results",
            "• Use WithVerticalAlignment(Top) for multi-line content",
            "• Combine with row height adjustment for optimal display",
            "• Text wrap works with all other formatting options",
            "• Available for both FluentCell and FluentRange classes"
        };

        for (int i = 0; i < tips.Length; i++)
        {
            worksheet.SetCell(49 + i, 1, tips[i])
                .WithFontColor(XLColor.Purple);
        }

        // Auto-fit rows to accommodate wrapped text
        worksheet.AutoFitRows();

        workbook.SaveAs("text-wrap-demo.xlsx");
        Console.WriteLine("   ✓ Created text-wrap-demo.xlsx showcasing text wrapping for cells and ranges!");
    }

    static void MultipleRangeSumDemo()
    {
        Console.WriteLine("\n13. Multiple Range SUM Demo - Showcasing enhanced WithSum capabilities");

        using var workbook = FluentWorkbook.Create();
        var worksheet = workbook.AddWorksheet("Multiple Range SUM");

        // Title
        worksheet.SetCell("A1", "Multiple Range SUM Formula Demonstration")
            .Bold()
            .WithFontSize(16)
            .WithFontColor(XLColor.DarkBlue);

        // Set up sample data
        worksheet.SetCell("A3", "Sample Data for SUM Operations").Bold().WithFontSize(14);
        
        // Revenue data
        worksheet.SetCell("A5", "Revenue:").Bold();
        worksheet.SetCell("B5", "Q1").Bold();
        worksheet.SetCell("C5", "Q2").Bold();
        worksheet.SetCell("D5", "Q3").Bold();
        worksheet.SetCell("E5", "Q4").Bold();
        
        worksheet.SetCell("A6", "Product A");
        worksheet.SetCell("B6", 1000);
        worksheet.SetCell("C6", 1200);
        worksheet.SetCell("D6", 1100);
        worksheet.SetCell("E6", 1300);
        
        worksheet.SetCell("A7", "Product B");
        worksheet.SetCell("B7", 800);
        worksheet.SetCell("C7", 900);
        worksheet.SetCell("D7", 950);
        worksheet.SetCell("E7", 1050);
        
        // Expenses data
        worksheet.SetCell("A9", "Expenses:").Bold();
        worksheet.SetCell("B9", "Q1").Bold();
        worksheet.SetCell("C9", "Q2").Bold();
        worksheet.SetCell("D9", "Q3").Bold();
        worksheet.SetCell("E9", "Q4").Bold();
        
        worksheet.SetCell("A10", "Marketing");
        worksheet.SetCell("B10", 200);
        worksheet.SetCell("C10", 250);
        worksheet.SetCell("D10", 220);
        worksheet.SetCell("E10", 280);
        
        worksheet.SetCell("A11", "Operations");
        worksheet.SetCell("B11", 300);
        worksheet.SetCell("C11", 320);
        worksheet.SetCell("D11", 310);
        worksheet.SetCell("E11", 340);

        // Single cell data
        worksheet.SetCell("A13", "Bonus:").Bold();
        worksheet.SetCell("B13", 500);

        // Example 1: Traditional single range SUM (FluentCell)
        worksheet.SetCell("A15", "EXAMPLES - FluentCell WithSum Methods:").Bold().WithFontSize(14).WithFontColor(XLColor.DarkGreen);
        
        worksheet.SetCell("A17", "1. Traditional single range SUM:");
        worksheet.SetCell("B17", "").WithSum("B6:E6").AsCurrency();
        worksheet.SetCell("C17", "=SUM(B6:E6) - Product A Total Revenue").WithFontColor(XLColor.Gray);

        // Example 2: Multiple ranges with FluentCell
        worksheet.SetCell("A19", "2. Multiple ranges SUM:");
        worksheet.SetCell("B19", "").WithSum("B6:E6", "B7:E7").AsCurrency();
        worksheet.SetCell("C19", "=SUM(B6:E6)+SUM(B7:E7) - Total Revenue").WithFontColor(XLColor.Gray);

        // Example 3: Mixed ranges and cells with FluentCell
        worksheet.SetCell("A21", "3. Mixed ranges and single cells:");
        worksheet.SetCell("B21", "").WithSum("B6:E6", "B13", "B7:E7").AsCurrency();
        worksheet.SetCell("C21", "=SUM(B6:E6)+B13+SUM(B7:E7) - Revenue + Bonus").WithFontColor(XLColor.Gray);

        // Example 4: With negative ranges (profit calculation) using FluentCell
        worksheet.SetCell("A23", "4. Profit calculation (Revenue - Expenses):");
        worksheet.SetCell("B23", "").WithSum("B6:E6", "B7:E7", "-B10:E10", "-B11:E11").AsCurrency();
        worksheet.SetCell("C23", "=SUM(B6:E6)+SUM(B7:E7)-SUM(B10:E10)-SUM(B11:E11)").WithFontColor(XLColor.Gray);

        // Example 5: Complex calculation with individual cells and ranges using FluentCell
        worksheet.SetCell("A25", "5. Complex: Revenue + Bonus - Marketing:");
        worksheet.SetCell("B25", "").WithSum("B6:E6", "B7:E7", "B13", "-B10:E10").AsCurrency();
        worksheet.SetCell("C25", "=SUM(B6:E6)+SUM(B7:E7)+B13-SUM(B10:E10)").WithFontColor(XLColor.Gray);

        // Example 6: Single negative cell using FluentCell
        worksheet.SetCell("A27", "6. Revenue minus single expense cell:");
        worksheet.SetCell("B27", "").WithSum("B6:E6", "-B10").AsCurrency();
        worksheet.SetCell("C27", "=SUM(B6:E6)-B10").WithFontColor(XLColor.Gray);

        // FluentRange examples
        worksheet.SetCell("A29", "EXAMPLES - FluentRange WithSum Methods:").Bold().WithFontSize(14).WithFontColor(XLColor.Purple);

        // Example 7: FluentRange single SUM
        worksheet.SetCell("A31", "7. FluentRange single SUM:");
        var range7 = worksheet.GetRange("B31:C31");
        range7.WithSum("B7:E7").AsCurrency().Merge();
        worksheet.SetCell("D31", "Range merged with =SUM(B7:E7)").WithFontColor(XLColor.Gray);

        // Example 8: FluentRange multiple ranges
        worksheet.SetCell("A33", "8. FluentRange multiple ranges:");
        var range8 = worksheet.GetRange("B33:C33");
        range8.WithSum("B6:E6", "B7:E7", "-B10:E10").AsCurrency().Merge();
        worksheet.SetCell("D33", "Range with complex SUM formula").WithFontColor(XLColor.Gray);

        // Example 9: FluentRange with single cells and ranges
        worksheet.SetCell("A35", "9. FluentRange mixed calculation:");
        var range9 = worksheet.GetRange("B35:C35");
        range9.WithSum("B6:E6", "B13", "-B10", "-B11").AsCurrency().Merge();
        worksheet.SetCell("D35", "Range with mixed cell types").WithFontColor(XLColor.Gray);

        // Summary section
        worksheet.SetCell("A37", "ENHANCED WithSum() CAPABILITIES SUMMARY:").Bold().WithFontSize(16).WithFontColor(XLColor.DarkBlue);

        var summaryData = new string[,]
        {
            { "Feature", "Example Syntax", "Generated Formula" },
            { "Single Range", ".WithSum(\"A1:A5\")", "=SUM(A1:A5)" },
            { "Multiple Ranges", ".WithSum(\"A1:A5\", \"B1:B5\")", "=SUM(A1:A5)+SUM(B1:B5)" },
            { "Negative Range", ".WithSum(\"A1:A5\", \"-B1:B5\")", "=SUM(A1:A5)-SUM(B1:B5)" },
            { "Single Cell", ".WithSum(\"A1\", \"B2\")", "=A1+B2" },
            { "Negative Cell", ".WithSum(\"A1\", \"-B2\")", "=A1-B2" },
            { "Mixed Types", ".WithSum(\"A1:A5\", \"B1\", \"-C1:C5\")", "=SUM(A1:A5)+B1-SUM(C1:C5)" },
            { "Complex Calc", ".WithSum(\"Rev1:Rev4\", \"Bonus\", \"-Exp1:Exp4\")", "=SUM(Rev1:Rev4)+Bonus-SUM(Exp1:Exp4)" }
        };

        var summaryRange = worksheet.GetRange("A39:C46");
        summaryRange.WithValues(summaryData)
                    .WithTableBorders()
                    .WithVerticalAlignment(XLAlignmentVerticalValues.Top);

        // Format summary header
        var summaryHeaderRange = worksheet.GetRange("A39:C39");
        summaryHeaderRange.Bold()
                          .WithBackgroundColor(XLColor.DarkBlue)
                          .WithFontColor(XLColor.White);

        // Usage tips
        worksheet.SetCell("A48", "USAGE TIPS:").Bold().WithFontSize(12).WithFontColor(XLColor.Orange);
        var tips = new[]
        {
            "• Use '-' prefix for any range or cell to subtract it: '-B1:B5', '-C10'",
            "• Single cells are referenced directly, ranges are wrapped in SUM()",
            "• Works with both FluentCell and FluentRange classes",
            "• Automatically handles complex formula generation",
            "• Perfect for financial calculations and data aggregation",
            "• Supports mixed single cells and ranges in one formula"
        };

        for (int i = 0; i < tips.Length; i++)
        {
            worksheet.SetCell(49 + i, 1, tips[i]).WithFontColor(XLColor.Orange);
        }

        // Format data sections
        var revenueHeaderRange = worksheet.GetRange("A5:E5");
        revenueHeaderRange.WithBackgroundColor(XLColor.LightGreen).Bold();

        var expenseHeaderRange = worksheet.GetRange("A9:E9");
        expenseHeaderRange.WithBackgroundColor(XLColor.LightPink).Bold();

        var revenueDataRange = worksheet.GetRange("A6:E7");
        revenueDataRange.WithBorder().AsCurrency();

        var expenseDataRange = worksheet.GetRange("A10:E11");
        expenseDataRange.WithBorder().AsCurrency();

        worksheet.SetCell("B13", 500).AsCurrency().WithBackgroundColor(XLColor.LightYellow);

        // Highlight the formula results
        var resultsRange = worksheet.GetRange("B17:B27");
        resultsRange.WithBackgroundColor(XLColor.LightCyan);

        worksheet.AutoFitColumns();

        workbook.SaveAs("multiple-range-sum-demo.xlsx");
        Console.WriteLine("   ✓ Created multiple-range-sum-demo.xlsx demonstrating enhanced WithSum() capabilities!");
    }
}

// Sample data classes
public class Employee
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public string Department { get; set; } = string.Empty;
    public decimal Salary { get; set; }
    public DateTime HireDate { get; set; }
}

public class FinancialRecord
{
    public string Account { get; set; } = string.Empty;
    public decimal Q1 { get; set; }
    public decimal Q2 { get; set; }
    public decimal Q3 { get; set; }
    public decimal Q4 { get; set; }
}

public class SaleRecord
{
    public string Product { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
    public decimal Price { get; set; }
    public int Quantity { get; set; }
}
