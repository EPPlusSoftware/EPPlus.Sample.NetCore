## This sample project is for EPPlus 6, a newer sample for EPPlus 7 is available here... [C#](https://github.com/EPPlusSoftware/EPPlus.Samples.CSharp) or [VB](https://github.com/EPPlusSoftware/EPPlus.Samples.VB)
## EPPlus samples

### EPPlus samples for .Net Core

The solution can be opened in Visual Studio for Windows or MacOS. On other operating systems please use...

```
dotnet restore
dotnet run
```

... to execute the samples.

|No|Sample|Description|
|---|---|-----------------|
|01|[Getting started](/01-GettingStarted/)|Basic usage of EPPlus: create a workbook, fill with data and some basic styling|
|02|[Read workbook](/02-ReadWorkbook/)|Read data from a workbook|
|03|[Async/Await](/03-UsingAsyncAwait/)|Using async/await methods for loading and saving data|
|04|[Loading data](/04-LoadingData/)|Load data into a worksheet from various types of objects and create a table.  It also demonstrates the Autofit columns feature.|
|05|[Import and Export csv files and create charts](/05-ImportAndExportCsvFiles/)|This sample shows how to load and save CSV files using the LoadFromText and SaveToText methods, how to use tables and how to use charts with more than one charttype and secondary axis.|
|06|[Calculate formulas](/06-FormulaCalculation/)|How to calculate formulas and add custom/missing functions in a workbook|
|07|[Open workbook and add data/chart](/07-OpenWorkbookAddDataAndChart/)|Opens an existing workbook, adds some data and a pie chart.|
|08|[Sales report](/08-SalesReport/)|Create a report with data from a SQL database.|
|09|[Performance and protection](/09-PerformanceAndProtection/)|Loads 65 000 rows, styles them and sets a password.|
|10|[Read data using Linq](/10-ReadDataUsingLinq/)|This sample shows how to use Linq with the Cells collection to read sample 9.|
|11|[Conditional formatting](/11-ConditionalFormatting/)|Demonstrates conditional formatting.|
|12|[Data validation](/12-DataValidation/)|How to add various types of data validation to a workbook and read existing validations.|
|13|[Filters](/13-Filter/)|How to apply filters in a worksheet, table or a pivot table|
|14|[Shapes & Images](/14-ShapesAndImages/)|Shows how to add shapes and format them in different ways.
|15|[Chart Styling & Themes ](/15-ChartsAndThemes/)|Load a theme and create various charts and style them.
|16|[Sparklines](/16-Sparklines/)|Demonstrates sparklines functionality.|
|17|[FX report](/17-FXReportFromDatabase/)|Exchange rates report with data from a SQL database. Demonstrates some of the chart capabilities of EPPlus|
|18|[Pivot tables](/18-PivotTables/)|Demonstrates the Pivot table functionality of EPPlus.|
|19|[Encryption and protection](/19-EncryptionAndProtection/)|This sample produces a quiz, where the template workbook is encrypted and password protected.|
|20|[Create filesystem report](/20-CreateFileSystemReport/)|Demonstrates usage of styling, printer settings, rich text, pie-, doughnut- and bar-charts, freeze panes and row/column outlines|
|21|[VBA - Visual Basic for Applications](/21-VBA/)|Demonstrates EPPlus support for VBA, includes a battleship game|
|22|[Ignore errors](/22-IgnoreErrors/)|Various samples on how to ignore error on cells.|
|23|[Comments](/23-Comments/)|Sample showing how to add notes and threaded comments.|
|24|[Slicers](/24-Slicers/)|Sample showing how to add Pivot Table slicers and Tabel slicers
|25|[Export to/from DataTable](/25-ImportAndExportDataTable)|Sample showing import/export rangedata with System.Data.DataTable
|26|[Form controls](/26-FormControls)|Sample showing how to add differnt form controls and how to group drawings.
|27|[Custom styles for tables and slicers](/27-CustomNamedStyles)|Sample showing how to create custom styles from tables, pivot tables and slicers.
|28|[Tables](/28-Tables)|Sample showing how to work with tables.
|29|[External links](/29-ExternalLinks)|Shows how to work with links to external workbooks
|30|[Sorting Ranges](/30-WorkingWithRanges)|Shows how to work with the Sort method for ranges and tables
|31|[Html Export](/31-HtmlExport)|Shows how to export tables and ranges to HTML
|32|[Json Export](/32-JsonExport)|Shows how to export tables and ranges to JSON
|33|[ToCollection](/33-ToCollection)|Shows how to export data from worksheets and tables into an IEnumerable&lt;T&gt; where T is a class.

### Output files
The samples above produces some workbooks - the name of each workbook indicates which sample that generated it. These workbooks are located in a subdirectory - named "SampleApp" - to the output directory of the sample project.

Also see wiki on https://github.com/EPPlusSoftware/EPPlus/wiki for more details
