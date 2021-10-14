using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Filter;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Reflection.Metadata;
using System.Text;

namespace EPPlusSamples
{
    public class SlicerSample
    {
        public static void Run(string connectionString)
        {
            using (var p = new ExcelPackage())
            {
                //Creates a worksheet with one table and several slicers.
                TableSlicerSample(p, connectionString);

                //Creates the source data for the pivot tables in a separate sheet.
                CreatePivotTableSourceWorksheet(p, connectionString);

                //Create a pivot table with a slicer connected to one field.
                PivotTableSlicerSample(p);
                //Create three slicers and two pivot tables, one that connects to both tables and two that connect to each of the tables.
                PivotTableOneSlicerToMultiplePivotTables(p);

                p.SaveAs(FileUtil.GetCleanFileInfo("24-Slicers.xlsx"));
            }
        }

        private static void PivotTableOneSlicerToMultiplePivotTables(ExcelPackage p)
        {
            var wsSource = p.Workbook.Worksheets["PivotTableSourceData"];
            var wsPivot = p.Workbook.Worksheets.Add("OneSlicerToTwoPivotTables");
            
            var pivotTable1 = wsPivot.PivotTables.Add(wsPivot.Cells["A15"], wsSource.Cells[wsSource.Dimension.Address], "PivotTable1");
            pivotTable1.RowFields.Add(pivotTable1.Fields["Company Name"]);
            pivotTable1.DataFields.Add(pivotTable1.Fields["Order Value"]);
            pivotTable1.DataFields.Add(pivotTable1.Fields["Tax"]);
            pivotTable1.DataFields.Add(pivotTable1.Fields["Freight"]);
            pivotTable1.DataOnRows = false;

            //To connect a slicer to multiple pivot tables the tables need to use the same pivot table cache, so we use pivotTable1's cache as source to pivotTable2...
            var pivotTable2 = wsPivot.PivotTables.Add(wsPivot.Cells["F15"], pivotTable1.CacheDefinition, "PivotTable2");
            pivotTable2.RowFields.Add(pivotTable2.Fields["Country"]);
            pivotTable2.DataFields.Add(pivotTable2.Fields["Order Value"]);
            pivotTable2.DataFields.Add(pivotTable2.Fields["Tax"]);
            pivotTable2.DataFields.Add(pivotTable2.Fields["Freight"]);
            pivotTable2.DataOnRows = false;

            var slicer1 = pivotTable1.Fields["Country"].AddSlicer();
            slicer1.Caption = "Country - Both";
            
            //Now add the second pivot table to the slicer cache. This require that the pivot tables share the same cache. 
            slicer1.Cache.PivotTables.Add(pivotTable2);
            slicer1.SetPosition(0, 0, 0, 0);
            slicer1.Style = eSlicerStyle.Light4;

            var slicer2 = pivotTable1.Fields["Company Name"].AddSlicer();
            slicer2.Caption = "Company Name - PivotTable1";
            slicer2.ChangeCellAnchor(eEditAs.Absolute);
            slicer2.SetPosition(0, 192);
            slicer2.SetSize(256, 260);

            var slicer3 = pivotTable2.Fields["Order date"].AddSlicer();
            slicer3.Caption = "Order date - PivotTable2";
            slicer3.ChangeCellAnchor(eEditAs.Absolute);
            slicer3.SetPosition(0, 448);
            slicer3.SetSize(256, 260);
        }
        private static void TableSlicerSample(ExcelPackage p, string connectionString)
        {
            var worksheet1 = p.Workbook.Worksheets.Add("TableSlicer");
            var worksheet2 = p.Workbook.Worksheets.Add("TableSlicerToOtherWorksheet");
            // Lets connect to the sample database for some data
            using (var sqlConn = new SQLiteConnection(connectionString))
            {
                sqlConn.Open();
                using (var sqlCmd = new SQLiteCommand("select CompanyName as 'Company Name', [Name] as Name, Email as 'E-Mail', c.Country as Country, orderdate as 'Order Date', (ordervalue) as 'Order Value',tax as Tax, freight As Freight, currency As Currency from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId ORDER BY 1,2 desc", sqlConn))
                {
                    using (var sqlReader = sqlCmd.ExecuteReader())
                    {
                        var range = worksheet1.Cells["A14"].LoadFromDataReader(sqlReader, true, "tblSalesData", OfficeOpenXml.Table.TableStyles.Medium6);
                        var tbl = worksheet1.Tables.GetFromRange(range);
                        range.Offset(1, 4, range.Rows - 1, 1).Style.Numberformat.Format = "yyyy-MM-dd";
                        range.Offset(1, 5, range.Rows - 1, 3).Style.Numberformat.Format = "#,##0";
                        range.AutoFitColumns();

                        //You can either add a slicer via the table column...
                        var slicer1 = tbl.Columns[0].AddSlicer();
                        //Filter values are compared to the Text property of the cell. 
                        slicer1.FilterValues.Add("Barton-Veum");
                        slicer1.FilterValues.Add("Christiansen LLC");
                        slicer1.SetPosition(0, 0, 0, 0);

                        //... or you can add it via the drawings collection.
                        var slicer2 = worksheet1.Drawings.AddTableSlicer(tbl.Columns["Country"]);
                        slicer2.SetPosition(0,192);

                        //A slicer also supports date groups...
                        var slicer3 = tbl.Columns["Order Date"].AddSlicer();
                        slicer3.FilterValues.Add(new ExcelFilterDateGroupItem(2017, 6));
                        slicer3.FilterValues.Add(new ExcelFilterDateGroupItem(2017, 7));
                        slicer3.FilterValues.Add(new ExcelFilterDateGroupItem(2017, 8));
                        slicer3.SetPosition(0, 384);

                        //... You can also add a slicer to another worksheet, if you use the drawings collection...
                        var slicer4 = worksheet2.Drawings.AddTableSlicer(tbl.Columns["E-Mail"]);
                        slicer4.Caption = "E-Mail - TableSlicer Worksheet";
                        slicer4.Description = "This slicer reference a table in another worksheet.";
                        slicer4.SetPosition(1, 0, 1, 0);
                        slicer4.To.Column = 7;  //Set the end position anchor to column H, to make the slicer wider.

                        var shape = worksheet2.Drawings.AddShape("InfoText", eShapeStyle.Rect);
                        shape.SetPosition(1, 0, 8, 0);
                        shape.Text = "This slicer filters the table located in the TableSlicer worksheet";
                    }
                }
                sqlConn.Close();
            }
            worksheet1.View.FreezePanes(14, 1);
        }        
        private static void PivotTableSlicerSample(ExcelPackage p)
        {
            var wsSource = p.Workbook.Worksheets["PivotTableSourceData"];
            var wsPivot = p.Workbook.Worksheets.Add("PivotTableSlicer");

            var pivotTable = wsPivot.PivotTables.Add(wsPivot.Cells["A1"], wsSource.Cells[wsSource.Dimension.Address], "PivotTable1");
            pivotTable.RowFields.Add(pivotTable.Fields["Company Name"]);
            pivotTable.DataFields.Add(pivotTable.Fields["Order Value"]);
            pivotTable.DataFields.Add(pivotTable.Fields["Tax"]);
            pivotTable.DataFields.Add(pivotTable.Fields["Freight"]);
            pivotTable.DataOnRows = false;

            var slicer1 = pivotTable.Fields["Name"].AddSlicer();
            slicer1.SetPosition(0, 0, 10, 0);
            slicer1.SetSize(400, 208);
            slicer1.Style = eSlicerStyle.Light4;
            slicer1.Cache.Data.Items.GetByValue("Brown Kutch").Hidden = true;
            slicer1.Cache.Data.Items.GetByValue("Tierra Ratke").Hidden = true;
            slicer1.Cache.Data.Items.GetByValue("Jamarcus Schimmel").Hidden = true;

            //Add a column with two columns and start showing the item 3.
            slicer1.ColumnCount = 2; //Use two columns on this slicer
            slicer1.StartItem = 3;   //First visible item is 3
            slicer1.Cache.Data.CrossFilter = eCrossFilter.ShowItemsWithNoData;
            slicer1.Cache.Data.SortOrder = eSortOrder.Descending;
        }
        private static void CreatePivotTableSourceWorksheet(ExcelPackage p, string connectionString)
        {
            var wsSource = p.Workbook.Worksheets.Add("PivotTableSourceData");
            //Lets connect to the sample database for some data
            using (var sqlConn = new SQLiteConnection(connectionString))
            {
                sqlConn.Open();
                using (var sqlCmd = new SQLiteCommand("select CompanyName as 'Company Name', [Name] as Name, Email as 'EMail', c.Country as Country, orderdate as 'Order Date', ordervalue as 'Order Value',tax as Tax, freight As Freight, currency As Currency from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId ORDER BY 1,2 desc", sqlConn))
                {
                    using (var sqlReader = sqlCmd.ExecuteReader())
                    {
                        var range = wsSource.Cells["A1"].LoadFromDataReader(sqlReader, true);
                        range.Offset(0, 0, 1, range.Columns).Style.Font.Bold = true;
                        range.Offset(1, 4, range.Rows - 1, 1).Style.Numberformat.Format = "yyyy-MM-dd hh:mm";
                        range.Offset(1, 5, range.Rows - 1, 3).Style.Numberformat.Format = "#,##0";
                    }
                }
                sqlConn.Close();
            }
            wsSource.Cells.AutoFitColumns();
        }
    }
}
