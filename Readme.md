# EPPlus samples

###EPPlus samples for .Net Core

Solution can be opened in Visual Studio for Windows or MacOS. On other operation systems please use...

```
dotnet restore
dotnet run
```

... to execute the samples.

|No|Sample|Description|
|---|---|-----------------|
|01|[Getting started](/EPPlus.Sample.NetCore/01-GettingStarted/)|Basic usage of EPPlus: create a workbook, fill with data and some basic styling|
|02|[Read workbook](/EPPlus.Sample.NetCore/02-ReadWorkbook/)|Read data from a workbook|
|03|[Async/Await](/EPPlus.Sample.NetCore/03-UsingAsyncAwait/)|Using async/await methods for loading and saving data|
|04|[Loading data](/EPPlus.Sample.NetCore/04-LoadingDataWithTables/)|Load data into a worksheet from various types of objects and create a table.  It also demonstrates the Autofit columns feature.|
|05|[Import and Export csv files and create charts](/EPPlus.Sample.NetCore/05-ImportAndExportCsvFiles/)|This sample shows how to load and save CSV files using the LoadFromText and SaveToText methods, how to use tables and how to use charts with more than one charttype and secondary axis.|
|06|[Calculate formulas](/EPPlus.Sample.NetCore/06-FormulaCalculation/)|How to calculate formulas and add custom/missing functions in a workbook|
|07|[Open workbook and add data/chart](/EPPlus.Sample.NetCore/07-OpenWorkbookAddDataAndChart/)|Opens an existing workbook, adds some data and a pie chart.|
|08|[Sales report](/EPPlus.Sample.NetCore/08-SalesReport/)|Create a report with data from a SQL database.|
|09|[Performance and protection](/EPPlus.Sample.NetCore/09-PerformanceAndProtection/)|Loads 65 000 rows, styles them and sets a password.|
|10|[Read data using Linq](/EPPlus.Sample.NetCore/10-ReadDataUsingLinq/)|This sample shows how to use Linq with the Cells collection to read sample 9.|
|11|[Conditional formatting](/EPPlus.Sample.NetCore/11-ConditionalFormatting/)|Demonstrates conditional formatting.|
|12|[Data validation](/EPPlus.Sample.NetCore/12-DataValidation/)|How to add various types of data validation to a workbook and read existing validations.|
|13|[Filters](/EPPlus.Sample.NetCore/13-Filter/)|How to apply filters in a worksheet or a table|
|14|[Shapes & Images](/EPPlus.Sample.NetCore/14-ShapesAndImages/)|Shows how to add shapes and format them in different ways.
|15|[Chart Styling & Themes ](/EPPlus.Sample.NetCore/15-ChartsAndThemes/)|Load a theme and create various charts and style them.
|16|[Sparklines](/EPPlus.Sample.NetCore/16-Sparklines/)|Demonstrates sparklines functionality.|
|17|[FX report](/EPPlus.Sample.NetCore/17-FXReportFromDatabase/)|Exchange rates report with data from a SQL database. Demonstrates some of the chart capabilities of EPPlus|
|18|[Pivot tables](/EPPlus.Sample.NetCore/18-PivotTables/)|Demonstrates the Pivot table functionality of EPPlus.|
|19|[Encryption and protection](/EPPlus.Sample.NetCore/19-EncryptionAndProtection/)|This sample produces a quiz, where the template workbook is encrypted and password protected.|
|20|[Create filesystem report](/EPPlus.Sample.NetCore/20-CreateFileSystemReport/)|Demonstrates usage of styling, printer settings, rich text, pie-, doughnut- and bar-charts, freeze panes|
|21|[VBA - Visual Basic for Applications](/EPPlus.Sample.NetCore/21-VBA/)|Demonstrates EPPlus support for VBA, includes a battleship game|

### Output files
The samples above produces some workbooks - the name of each workbook indicates which sample that generated it. These workbooks are located in a subdirectory - named "SampleApp" - to the output directory of the sample project.


### Non windows operating systems.
Non-windows operating systems will requires libgdiplus to be installed. 
Please use your favorite package manager to install it. 
For example:

Homebrew on MacOS:
```
brew install mono-libgdiplus
```

apt-get:
```
apt-get install libgdiplus
```

Also see wiki on https://github.com/JanKallman/EPPlus/wiki for more details
