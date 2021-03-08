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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Data.SqlClient;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using System.Data.SQLite;

namespace EPPlusSamples.PivotTables
{
    /// <summary>
    /// This class shows how to use pivottables 
    /// </summary>
    public static class PivotTablesSample
    {
        public class SalesDTO
        {
            public string CompanyName { get; set; }            
            public string Name { get; set; }
            public string Email { get; set; }
            public string Country { get; set; }
            public int OrderId { get; set; }
            public DateTime OrderDate { get; set; }
            public decimal OrderValue { get; set; }
            public decimal Tax { get; set; }
            public decimal Freight { get; set; }
            public string Currency { get; set; }
        }
        public static string Run(string connectionStr)
        {
            var list = GetDataFromSQL(connectionStr);

            FileInfo newFile = FileOutputUtil.GetFileInfo("18-PivotTables.xlsx");
            using (ExcelPackage pck = new ExcelPackage(newFile))
            {
                // get the handle to the existing worksheet
                var wsData = pck.Workbook.Worksheets.Add("SalesData");

                var dataRange = wsData.Cells["A1"].LoadFromCollection
                    (
                    from s in list
                    orderby s.Name
                    select s,
                   true, OfficeOpenXml.Table.TableStyles.Medium2);

                wsData.Cells[2, 6, dataRange.End.Row, 6].Style.Numberformat.Format = "mm-dd-yy";
                wsData.Cells[2, 7, dataRange.End.Row, 11].Style.Numberformat.Format = "#,##0";

                dataRange.AutoFitColumns();

                var pt1 = CreatePivotTableWithPivotChart(pck, dataRange);
                var pt2 = CreatePivotTableWithDataGrouping(pck, dataRange);
                var pt3 = CreatePivotTableWithPageFilter(pck, pt2.CacheDefinition);
                var pt4 = CreatePivotTableWithASlicer(pck, pt2.CacheDefinition);
                var pt5 = CreatePivotTableWithACalculatedField(pck, pt2.CacheDefinition);

                //Filter samples
                var pt6 = CreatePivotTableCaptionFilter(pck, dataRange);

                pck.Save();
            }
            return newFile.FullName;
        }
        private static ExcelPivotTable CreatePivotTableWithPivotChart(ExcelPackage pck, ExcelRangeBase dataRange)
        {
            var wsPivot = pck.Workbook.Worksheets.Add("PivotSimple");
            var pivotTable = wsPivot.PivotTables.Add(wsPivot.Cells["A1"], dataRange, "PerCountry");

            pivotTable.RowFields.Add(pivotTable.Fields["Country"]);
            var dataField = pivotTable.DataFields.Add(pivotTable.Fields["OrderValue"]);
            dataField.Format = "#,##0";
            pivotTable.DataOnRows = true;

            var chart = wsPivot.Drawings.AddPieChart("PivotChart", ePieChartType.PieExploded3D, pivotTable);
            chart.SetPosition(1, 0, 4, 0);
            chart.SetSize(800, 600);
            chart.Legend.Remove();
            chart.Series[0].DataLabel.ShowCategory = true;
            chart.Series[0].DataLabel.Position = eLabelPosition.OutEnd;
            chart.StyleManager.SetChartStyle(ePresetChartStyle.Pie3dChartStyle6);
            return pivotTable;
        }

        private static ExcelPivotTable CreatePivotTableWithDataGrouping(ExcelPackage pck, ExcelRangeBase dataRange)
        {
            var wsPivot2 = pck.Workbook.Worksheets.Add("PivotDateGrp");
            var pivotTable2 = wsPivot2.PivotTables.Add(wsPivot2.Cells["A3"], dataRange, "PerEmploeeAndQuarter");

            pivotTable2.RowFields.Add(pivotTable2.Fields["Name"]);

            //Add a rowfield
            var rowField = pivotTable2.RowFields.Add(pivotTable2.Fields["OrderDate"]);
            //This is a date field so we want to group by Years and quaters. This will create one additional field for years.
            rowField.AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Quarters);
            //Get the Quaters field and change the texts
            var quaterField = pivotTable2.Fields.GetDateGroupField(eDateGroupBy.Quarters);
            quaterField.Items[0].Text = "<"; //Values below min date, but we use auto so its not used
            quaterField.Items[1].Text = "Q1";
            quaterField.Items[2].Text = "Q2";
            quaterField.Items[3].Text = "Q3";
            quaterField.Items[4].Text = "Q4";
            quaterField.Items[5].Text = ">"; //Values above max date, but we use auto so its not used

            //Add a pagefield
            var pageField = pivotTable2.PageFields.Add(pivotTable2.Fields["CompanyName"]);

            //Add the data fields and format them
            ExcelPivotTableDataField dataField;
            dataField = pivotTable2.DataFields.Add(pivotTable2.Fields["OrderValue"]);
            dataField.Format = "#,##0";
            dataField = pivotTable2.DataFields.Add(pivotTable2.Fields["Tax"]);
            dataField.Format = "#,##0";
            dataField = pivotTable2.DataFields.Add(pivotTable2.Fields["Freight"]);
            dataField.Format = "#,##0";

            //We want the datafields to appear in columns
            pivotTable2.DataOnRows = false;
            return pivotTable2;
        }
        private static ExcelPivotTable CreatePivotTableWithPageFilter(ExcelPackage pck, ExcelPivotCacheDefinition pivotCache)
        {
            var wsPivot3 = pck.Workbook.Worksheets.Add("PivotWithPageField");
            
            //Create a new pivot table using the same cache as pivot table 2.
            var pivotTable3 = wsPivot3.PivotTables.Add(wsPivot3.Cells["A3"], pivotCache, "PerEmploeeSelectedCompanies");

            pivotTable3.RowFields.Add(pivotTable3.Fields["Name"]);

            //Add a rowfield
            var rowField = pivotTable3.RowFields.Add(pivotTable3.Fields["OrderDate"]);

            //Add a pagefield
            var pageField = pivotTable3.PageFields.Add(pivotTable3.Fields["CompanyName"]);
            pageField.Items.Refresh();  //Refresh the items from the source range.
            
            pageField.Items[1].Hidden = true;   //Hide item with index 1 in the items collection
            pageField.Items.GetByValue("Walsh LLC").Hidden = true;  //Hide the item with supplied the value . 
            //pageField.Items.SelectSingleItem(3); //You can also select a single item with this method

            //Add the data fields and format them
            ExcelPivotTableDataField dataField;
            dataField = pivotTable3.DataFields.Add(pivotTable3.Fields["OrderValue"]);
            dataField.Format = "#,##0";
            dataField = pivotTable3.DataFields.Add(pivotTable3.Fields["Tax"]);
            dataField.Format = "#,##0";
            dataField = pivotTable3.DataFields.Add(pivotTable3.Fields["Freight"]);
            dataField.Format = "#,##0";

            
            //We want the datafields to appear in columns
            pivotTable3.DataOnRows = false;
            return pivotTable3;
        }
        private static ExcelPivotTable CreatePivotTableWithASlicer(ExcelPackage pck, ExcelPivotCacheDefinition pivotCache)
        {
            //This method connects a slicer to the pivot table. Also see sample 24 for more detailed samples on slicers.
            var wsPivot4 = pck.Workbook.Worksheets.Add("PivotWithSlicer");

            //Create a new pivot table using the same cache as pivot table 2.
            var pivotTable4 = wsPivot4.PivotTables.Add(wsPivot4.Cells["A3"], pivotCache, "PerEmploeeSelectedCompSlicer");

            pivotTable4.RowFields.Add(pivotTable4.Fields["Name"]);

            //Add a rowfield
            pivotTable4.RowFields.Add(pivotTable4.Fields["OrderDate"]);

            //Add slicer
            var companyNameField = pivotTable4.Fields["CompanyName"];
            var slicer = companyNameField.AddSlicer();
            slicer.SetPosition(3, 0, 5, 0); //Set top left to row 4, column F

            companyNameField.Items.Refresh();  //Refresh the items from the source range.

            companyNameField.Items[1].Hidden = true;   //Hide item with index 1 in the items collection
            companyNameField.Items.GetByValue("Walsh LLC").Hidden = true;  //Hide the item with supplied the value . 

            //Add the data fields and format them
            ExcelPivotTableDataField dataField;
            dataField = pivotTable4.DataFields.Add(pivotTable4.Fields["OrderValue"]);
            dataField.Format = "#,##0";
            dataField = pivotTable4.DataFields.Add(pivotTable4.Fields["Tax"]);
            dataField.Format = "#,##0";
            dataField = pivotTable4.DataFields.Add(pivotTable4.Fields["Freight"]);
            dataField.Format = "#,##0";

            //We want the data fields to appear in columns
            pivotTable4.DataOnRows = false;
            return pivotTable4;
        }
        private static ExcelPivotTable CreatePivotTableWithACalculatedField(ExcelPackage pck, ExcelPivotCacheDefinition pivotCache)
        {
            //This method connects a slicer to the pivot table. Also see sample 24 for more detailed samples on slicers.
            var wsPivot4 = pck.Workbook.Worksheets.Add("PivotWithCalculatedField");

            //Create a new pivot table using the same cache as pivot table 2.
            var pivotTable4 = wsPivot4.PivotTables.Add(wsPivot4.Cells["A3"], pivotCache, "PerWithCalculatedField");

            pivotTable4.RowFields.Add(pivotTable4.Fields["CompanyName"]);
            //Be careful with formulas as they are not validated and can cause the pivot table to become corrupt. 

            //Be careful with formulas as they can cause the pivot table to become corrupt if they are entered invalidly.
            var calcField = pivotTable4.Fields.AddCalculatedField("Total", "'OrderValue'+'Tax'+'Freight'");
            calcField.Format = "#,##0";

            //Add the data fields and format them
            ExcelPivotTableDataField dataField;
            dataField = pivotTable4.DataFields.Add(pivotTable4.Fields["OrderValue"]);
            dataField.Format = "#,##0";
            dataField = pivotTable4.DataFields.Add(pivotTable4.Fields["Tax"]);
            dataField.Format = "#,##0";
            dataField = pivotTable4.DataFields.Add(pivotTable4.Fields["Freight"]);
            dataField.Format = "#,##0";
            dataField = pivotTable4.DataFields.Add(pivotTable4.Fields["Total"]);
            dataField.Format = "#,##0";


            //We want the data fields to appear in columns
            pivotTable4.DataOnRows = false;
            return pivotTable4;
        }
        private static ExcelPivotTable CreatePivotTableCaptionFilter(ExcelPackage pck, ExcelRangeBase dataRange)
        {
            var wsPivot4 = pck.Workbook.Worksheets.Add("PivotWithCaptionFilter");

            //Create a new pivot table with a new cache.
            var pivotTable4 = wsPivot4.PivotTables.Add(wsPivot4.Cells["A3"], dataRange, "WithCaptionFilter");

            var rowField1 = pivotTable4.RowFields.Add(pivotTable4.Fields["Name"]);
            
            //Add the Caption filter (Label filter in Excel) to the pivot table.
            rowField1.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotBeginsWith, "C");
            
            //Add a rowfield
            var rowField2 = pivotTable4.RowFields.Add(pivotTable4.Fields["OrderDate"]);

            //Add a date value filter to the pivot table.
            rowField2.Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateBetween, new DateTime(2017,8,1), new DateTime(2017, 8, 31));

            //Filters will apply on top of any selection made directly on the items.
            rowField2.Items.Refresh();
            rowField2.Items[8].Hidden = true;

            //Number formats can be set directly on fields as well as on datafields...
            pivotTable4.Fields["OrderDate"].Format = "yyyy-MM-dd hh:mm:ss";
            pivotTable4.Fields["OrderValue"].Format = "#,##0";
            pivotTable4.Fields["Tax"].Format = "#,##0";
            pivotTable4.Fields["Freight"].Format = "#,##0";

            //Add the data fields and format them
            pivotTable4.DataFields.Add(pivotTable4.Fields["OrderValue"]);
            pivotTable4.DataFields.Add(pivotTable4.Fields["Tax"]);
            pivotTable4.DataFields.Add(pivotTable4.Fields["Freight"]);

            //We want the datafields to appear in columns
            pivotTable4.DataOnRows = false;
            return pivotTable4;
        }
        private static List<SalesDTO> GetDataFromSQL(string connectionStr)
        {
            var ret = new List<SalesDTO>();
            using (var sqlConn = new SQLiteConnection(connectionStr))
            {
                sqlConn.Open();
                using (var sqlCmd = new SQLiteCommand("select companyName as CompanyName, [name] as Name, email as Email, c.country as Country, o.OrderId as OrderId, orderdate as OrderDate, ordervalue as OrderValue, tax as Tax, freight as Freight, currency Currency from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId ORDER BY OrderDate, OrderValue desc", sqlConn))
                {
                    using (var sqlReader = sqlCmd.ExecuteReader())
                    {
                        //Get the data and fill rows 5 onwards
                        while (sqlReader.Read())
                        {
                            ret.Add(new SalesDTO
                            {
                                CompanyName = sqlReader["companyName"].ToString(),
                                Name = sqlReader["name"].ToString(),
                                Email = sqlReader["email"].ToString(),
                                Country = sqlReader["country"].ToString(),
                                OrderId = Convert.ToInt32(sqlReader["orderId"]),
                                OrderDate = (DateTime)sqlReader["OrderDate"],
                                OrderValue = Convert.ToDecimal(sqlReader["OrderValue"]),
                                Tax = Convert.ToDecimal(sqlReader["tax"]),
                                Freight = Convert.ToDecimal(sqlReader["freight"]),
                                Currency = sqlReader["currency"].ToString(),
                            });
                        }
                    }
                }
            }
            return ret;
        }
    }
}