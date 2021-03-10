using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Table;
using System;
using System.Data.SQLite;
using System.Drawing;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    /// <summary>
    /// This sample demonstrates how work with Excel tables in EPPlus.
    /// Tables can easily be added by many of the ExcelRange - Load methods as demonstrated in earlier sample.
    /// This sample will focus on how to add and setup tables from the ExcelWorksheet.Tables collection.
    /// </summary>
    public static class TablesSample
    {
        public static async Task RunAsync(string connectionString)
        {
            using (var p = new ExcelPackage())
            {
                await CreateTableWithACalculatedColumnAsync(connectionString, p).ConfigureAwait(false);
                await StyleTablesAsync(connectionString, p).ConfigureAwait(false);
                await CreateTableFilterAndSlicerAsync(connectionString, p).ConfigureAwait(false);

                p.SaveAs(FileOutputUtil.GetFileInfo("28-Tables.xlsx"));
            }
        }
        /// <summary>
        /// This sample creates a table with a calculated column. A totals row is added and styling is applied to some of the columns.
        /// </summary>
        /// <param name="connectionString">The connection string to the database</param>
        /// <param name="p">The package</param>
        /// <returns></returns>
        private static async Task CreateTableWithACalculatedColumnAsync(string connectionString, ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("SimpleTable");

            var range = await LoadDataAsync(connectionString, ws).ConfigureAwait(false);
            var tbl=ws.Tables.Add(range, "Table1");
            
            tbl.ShowTotal = true;
            //Format the OrderDate column and add a Count Numbers subtotal.
            tbl.Columns["OrderDate"].TotalsRowFunction = RowFunctions.CountNums;
            tbl.Columns["OrderDate"].DataStyle.NumberFormat.Format = "yyyy-MM-dd";
            tbl.Columns["OrderDate"].TotalsRowStyle.NumberFormat.Format = "#,##0";

            //Format the OrderValue column and add a Sum subtotal.
            tbl.Columns["OrderValue"].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns["OrderValue"].DataStyle.NumberFormat.Format = "#,##0";
            tbl.Columns["OrderValue"].TotalsRowStyle.NumberFormat.Format = "#,##0";

            //Add a calculated formula referencing the OrderValue column within the same row.
            tbl.Columns.Add(1);
            var addedcolumn = tbl.Columns[tbl.Columns.Count - 1];
            addedcolumn.Name = "OrderValue with Tax";
            addedcolumn.CalculatedColumnFormula = "Table1[[#This Row],[OrderValue]] * 110%"; //Sets the calculated formula referencing the OrderValue column within this row.
            addedcolumn.TotalsRowFunction = RowFunctions.Sum;
            addedcolumn.DataStyle.NumberFormat.Format = "#,##0";
            addedcolumn.TotalsRowStyle.NumberFormat.Format = "#,##0";

            tbl.ShowLastColumn = true;            

            tbl.Range.AutoFitColumns();
        }
        /// <summary>
        /// This sample creates a two table and a custom table style. The first table is styled using different style objects of the table. 
        /// The second table is styled using the custom table style
        /// </summary>
        /// <param name="connectionString"></param>
        /// <param name="p"></param>
        /// <returns></returns>
        private static async Task StyleTablesAsync(string connectionString, ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("StyleTables");

            var range1 = await LoadDataAsync(connectionString, ws).ConfigureAwait(false);

            //Add the table and set some styles and properties.
            var tbl1 = ws.Tables.Add(range1, "StyleTable1");
            tbl1.TableStyle = TableStyles.Medium24;
            tbl1.DataStyle.Font.Size = 10;
            tbl1.Columns["E-Mail"].DataStyle.Font.Underline=OfficeOpenXml.Style.ExcelUnderLineType.Single;
            tbl1.HeaderRowStyle.Font.Italic = true;
            tbl1.ShowTotal = true;
            tbl1.TotalsRowStyle.Font.Italic = true;
            tbl1.Range.Style.Font.Name = "Arial";
            tbl1.Range.AutoFitColumns();
            
            //Add two rows at the end.
            var addedRange = tbl1.AddRow(2);
            addedRange.Offset(0, 0, 1, 1).Value = "Added Row 1";
            addedRange.Offset(1, 0, 1, 1).Value = "Added Row 2";

            //Add a custom formula to display number of items in the CompanyName column
            tbl1.Columns[0].TotalsRowFormula= "\"Total Count is \" & SUBTOTAL(103,StyleTable1[CompanyName])";
            tbl1.Columns[0].TotalsRowStyle.Font.Color.SetColor(Color.Red);

            //We create a custom named style via the Workbook.Styles object. For more samples on custom styles see sample 27
            var customStyleName = "EPPlus Created Style";
            var customStyle = p.Workbook.Styles.CreateTableStyle(customStyleName, TableStyles.Medium13);
            customStyle.HeaderRow.Style.Font.Color.SetColor(eThemeSchemeColor.Text1);
            customStyle.FirstColumn.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent5);
            customStyle.FirstColumn.Style.Fill.BackgroundColor.Tint=0.3;
            customStyle.FirstColumn.Style.Font.Color.SetColor(eThemeSchemeColor.Text1);

            var range2 = await LoadDataAsync(connectionString, ws, "K1").ConfigureAwait(false);
            var tbl2 = ws.Tables.Add(range2, "StyleTable2");            
            //To apply the custom style we set the StyleName property to the name we choose for our style.
            tbl2.StyleName = customStyleName;
            tbl2.ShowFirstColumn = true;

            tbl2.Range.AutoFitColumns();
        }
        /// <summary>
        /// This sample creates a table and a slicer. 
        /// </summary>
        /// <param name="connectionString">The connection string to the database</param>
        /// <param name="p">The package</param>
        /// <returns></returns>
        private static async Task CreateTableFilterAndSlicerAsync(string connectionString, ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("Slicer");

            var range = await LoadDataAsync(connectionString, ws).ConfigureAwait(false);
            var tbl = ws.Tables.Add(range, "FilterTable1");
            tbl.TableStyle = TableStyles.Medium1;
            
            //Add a slicer and filter on company name. A table slicer is connected to a table columns value filter.
            var slicer1 = tbl.Columns[0].AddSlicer();
            slicer1.FilterValues.Add("Cremin-Kihn");
            slicer1.FilterValues.Add("Senger LLC");
            range.AutoFitColumns();

            //Apply the column filter, otherwise the slicer may be hidden when the filter is applied.
            tbl.AutoFilter.ApplyFilter();
            slicer1.SetPosition(2, 0, 10, 0);

            //For more samples on filters and slicers see Sample 13 and 24.
        }
        private static async Task<ExcelRangeBase> LoadDataAsync(string connectionString, ExcelWorksheet ws, string startCell="A1")
        {
            ExcelRangeBase range;
            //Lets connect to the sample database for some data
            using (var sqlConn = new SQLiteConnection(connectionString))
            {
                sqlConn.Open();
                using (var sqlCmd = new SQLiteCommand("select companyname as CompanyName, [name] as [Name], Email as [E-Mail], c.Country as Country, o.orderid as OrderId, orderdate as OrderDate, ordervalue as OrderValue, currency as Currency from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId ORDER BY 1,2 desc", sqlConn))
                {
                    using (var sqlReader = sqlCmd.ExecuteReader())
                    {
                        range = await ws.Cells[startCell].LoadFromDataReaderAsync(sqlReader, true);
                    }
                }
                sqlConn.Close();
            }
            return range;
        }
    }
}
