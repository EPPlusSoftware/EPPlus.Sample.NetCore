/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB           Initial release EPPlus 5
 *************************************************************************************************/
using EPPlusSamples;
using OfficeOpenXml;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Data.SQLite;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace EPPlusSampleApp.Core
{
    public class Filter
    {
        public static async Task RunAsync(string connectionString)
        {
            var p = new ExcelPackage();

            //Autofilter on the worksheet
            await ValueFilter(connectionString, p);
            await DateTimeFilter(connectionString, p);
            await CustomFilter(connectionString, p);
            await Top10Filter(connectionString, p);
            await DynamicAboveAverageFilter(connectionString, p);
            await DynamicDateAugustFilter(connectionString, p);

            //Filter on a table, also see sample 24-Slicers. 
            await TableFilter(connectionString, p);
            
            //Filter on a pivot table, also see sample 24-Slicers. 
            await PivotTableFilter(connectionString, p);

            p.SaveAs(FileOutputUtil.GetFileInfo("13-Filters.xlsx"));
        }

        private static async Task ValueFilter(string connectionString, ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("ValueFilter");
            ExcelRangeBase range = await LoadFromDatabase(connectionString, ws);

            range.AutoFilter = true;
            var colCompany = ws.AutoFilter.Columns.AddValueFilterColumn(0);
            colCompany.Filters.Add("Walsh LLC");
            colCompany.Filters.Add("Harber-Goldner");
            ws.AutoFilter.ApplyFilter();
            range.AutoFitColumns(0);
        }
        private static async Task DateTimeFilter(string connectionString, ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("DateTimeFilter");
            ExcelRangeBase range = await LoadFromDatabase(connectionString, ws);

            range.AutoFilter = true;
            var col = ws.AutoFilter.Columns.AddValueFilterColumn(5);
            col.Filters.Add(new ExcelFilterDateGroupItem(2017, 8));
            col.Filters.Add(new ExcelFilterDateGroupItem(2017, 7, 5));
            col.Filters.Add(new ExcelFilterDateGroupItem(2017, 7, 7));
            ws.AutoFilter.ApplyFilter();
            range.AutoFitColumns(0);
        }
        private static async Task CustomFilter(string connectionString, ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("CustomFilter");
            ExcelRangeBase range = await LoadFromDatabase(connectionString, ws);

            range.AutoFilter = true;
            var colCompany = ws.AutoFilter.Columns.AddCustomFilterColumn(6);
            colCompany.And = true;
            colCompany.Filters.Add(new ExcelFilterCustomItem("999.99",eFilterOperator.GreaterThan));
            colCompany.Filters.Add(new ExcelFilterCustomItem("1500", eFilterOperator.LessThanOrEqual));
            ws.AutoFilter.ApplyFilter();
            range.AutoFitColumns(0);
        }
        private static async Task Top10Filter(string connectionString, ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("Top10Filter");
            ExcelRangeBase range = await LoadFromDatabase(connectionString, ws);

            range.AutoFilter = true;
            var colTop10 = ws.AutoFilter.Columns.AddTop10FilterColumn(6);
            colTop10.Percent = false;    //If set to true, the value takes top the percentage. Otherwise it relates to the number of items.
            colTop10.Value = 10;         //The value to relate to.
            colTop10.Top = false;        //Top if true, bottom if false
            ws.AutoFilter.ApplyFilter();
            range.AutoFitColumns(0);
        }
        private static async Task DynamicAboveAverageFilter(string connectionString, ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("DynamicAboveAverageFilter");
            ExcelRangeBase range = await LoadFromDatabase(connectionString, ws);

            range.AutoFilter = true;
            var colDynamic = ws.AutoFilter.Columns.AddDynamicFilterColumn(6);
            colDynamic.Type = eDynamicFilterType.AboveAverage;
            ws.AutoFilter.ApplyFilter();
            range.AutoFitColumns(0);
        }
        private static async Task DynamicDateAugustFilter(string connectionString, ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("DynamicAugustFilter");
            ExcelRangeBase range = await LoadFromDatabase(connectionString, ws);

            range.AutoFilter = true;
            var colDynamic = ws.AutoFilter.Columns.AddDynamicFilterColumn(5);
            colDynamic.Type = eDynamicFilterType.M8;
            ws.AutoFilter.ApplyFilter();
            range.AutoFitColumns(0);
        }
        private static async Task TableFilter(string connectionString, ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("TableFilter");
            ExcelRangeBase range = await LoadFromDatabase(connectionString, ws);

            var tbl = ws.Tables.Add(range, "tblFilter");
            tbl.TableStyle = OfficeOpenXml.Table.TableStyles.Medium23;
            tbl.ShowFilter = true;
            //Add a value filter
            var colCompany = tbl.AutoFilter.Columns.AddValueFilterColumn(0);
            colCompany.Filters.Add("Walsh LLC");
            colCompany.Filters.Add("Harber-Goldner");
            colCompany.Filters.Add("Sporer, Mertz and Jaskolski");

            //Add a second filter on order value
            var colOrderValue = tbl.AutoFilter.Columns.AddCustomFilterColumn(6);
            colOrderValue.Filters.Add(new ExcelFilterCustomItem("500", eFilterOperator.GreaterThanOrEqual));
            tbl.AutoFilter.ApplyFilter();
            range.AutoFitColumns(0);
        }
        private static async Task PivotTableFilter(string connectionString, ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("PivotTableFilter");
            ExcelRangeBase range = await LoadFromDatabase(connectionString, ws);

            var tbl = ws.Tables.Add(range, "ptFilter");
            tbl.TableStyle = OfficeOpenXml.Table.TableStyles.Medium23;

            var pt1=ws.PivotTables.Add(ws.Cells["J1"], tbl, "PivotTable1");
            var rowField = pt1.RowFields.Add(pt1.Fields["CompanyName"]);
            var dataField = pt1.DataFields.Add(pt1.Fields["OrderValue"]);

            //First deselect a company in the items list. To do so we first need to refresh the items from the range.
            rowField.Items.Refresh();  //Refresh the items from the range.
            rowField.Items.GetByValue("Sporer, Mertz and Jaskolski").Hidden=true;
            //Add a caption filter on Company Name between A and D
            rowField.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionBetween, "A", "D");
            //Add a value filter where OrderValue >= 100
            rowField.Filters.AddValueFilter(ePivotTableValueFilterType.ValueGreaterThanOrEqual, dataField, 100);

            //Add a second pivot table with some different filters.
            var pt2 = ws.PivotTables.Add(ws.Cells["M1"], tbl, "PivotTable2");
            var rowField1 = pt2.RowFields.Add(pt2.Fields["Currency"]);
            var rowField2 = pt2.RowFields.Add(pt2.Fields["OrderDate"]);
            rowField2.Format = "yyyy-MM-dd";
            var dataField1 = pt2.DataFields.Add(pt2.Fields["OrderValue"]);
            var dataField2 = pt2.DataFields.Add(pt2.Fields["OrderId"]);
            dataField2.Function = DataFieldFunctions.CountNums;

            var slicer = rowField1.AddSlicer();
            slicer.SetPosition(11, 0, 9, 0);
            //Add a date filter between first of Mars 2017 to 30th of June
            rowField2.Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateBetween, new DateTime(2017, 3, 1), new DateTime(2017, 6, 30));
            //Add a filter on the bottom 25 percent of the OrderValue
            rowField2.Filters.AddTop10Filter(ePivotTableTop10FilterType.Percent, dataField1, 25, false);
            pt2.DataOnRows = false;

            range.AutoFitColumns(0);
        }

        private static async Task<ExcelRangeBase> LoadFromDatabase(string connectionString, ExcelWorksheet ws)
        {
            ExcelRangeBase range;
            using (var sqlConn = new SQLiteConnection(connectionString))
            {
                sqlConn.Open();
                using (var sqlCmd = new SQLiteCommand("select companyName as CompanyName, [name] as Name, email as Email, c.country as Country, o.OrderId as OrderId, orderdate as OrderDate, ordervalue as OrderValue, currency Currency from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId ORDER BY 1,2 desc", sqlConn))
                {
                    using (var sqlReader = sqlCmd.ExecuteReader())
                    {
                        range = await ws.Cells["A1"].LoadFromDataReaderAsync(sqlReader, true);
                        range.Offset(0, 0, 1, range.Columns).Style.Font.Bold = true;
                        range.Offset(0, 5, range.Rows, 1).Style.Numberformat.Format = "yyyy-MM-dd";
                    }
                    //Set the numberformat
                }
            }

            return range;
        }
    }
}
