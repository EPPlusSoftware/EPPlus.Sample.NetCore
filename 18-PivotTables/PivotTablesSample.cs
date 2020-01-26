/*************************************************************************************************
  Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/24/2020         Jan Källman & Mats Alm       Initial release EPPlus 5
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

                var wsPivot = pck.Workbook.Worksheets.Add("PivotSimple");
                var pivotTable1 = wsPivot.PivotTables.Add(wsPivot.Cells["A1"], dataRange, "PerCountry");

                pivotTable1.RowFields.Add(pivotTable1.Fields["Country"]);
                var dataField = pivotTable1.DataFields.Add(pivotTable1.Fields["OrderValue"]);
                dataField.Format="#,##0";
                pivotTable1.DataOnRows = true;

                var chart = wsPivot.Drawings.AddPieChart("PivotChart", ePieChartType.PieExploded3D, pivotTable1);
                chart.SetPosition(1, 0, 4, 0);
                chart.SetSize(800, 600);
                chart.Legend.Remove();
                chart.Series[0].DataLabel.ShowCategory = true;
                chart.Series[0].DataLabel.Position = eLabelPosition.OutEnd;
                chart.StyleManager.SetChartStyle(ePresetChartStyle.Pie3dChartStyle6);
                    
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
                dataField = pivotTable2.DataFields.Add(pivotTable2.Fields["OrderValue"]);
                dataField.Format = "#,##0";
                dataField = pivotTable2.DataFields.Add(pivotTable2.Fields["Tax"]);
                dataField.Format = "#,##0";
                dataField = pivotTable2.DataFields.Add(pivotTable2.Fields["Freight"]);
                dataField.Format = "#,##0";
                
                //We want the datafields to appear in columns
                pivotTable2.DataOnRows = false;

                pck.Save();
            }
            return newFile.FullName;
        }

        //private static List<SalesDTO> GetRandomData()
        //{   
        //    List<SalesDTO> ret = new List<SalesDTO>();  
        //    var firstNames = new string[] {"John", "Gunnar", "Karl", "Alice"};
        //    var lastNames = new string[] {"Smith", "Johansson", "Lindeman"};
        //    Random r = new Random();
        //    for (int i = 0; i < 500; i++)
        //    {
        //        ret.Add(
        //            new SalesDTO()
        //            {
        //                FirstName = firstNames[r.Next(4)],
        //                LastName = lastNames[r.Next(3)],
        //                OrderDate = new DateTime(2002, 1, 1).AddDays(r.Next(1000)),
        //                Title="Sales Representative",
        //                SubTotal = r.Next(100, 10000),
        //                Tax = 0,
        //                Freight = 0
        //            });
        //    }
        //    return ret;
        //}

        private static List<SalesDTO> GetDataFromSQL(string connectionStr)
        {
            var ret = new List<SalesDTO>();
            using (var sqlConn = new SQLiteConnection(connectionStr))
            {
                sqlConn.Open();
                using (var sqlCmd = new SQLiteCommand("select companyName as CompanyName, [name] as Name, email as Email, country as Country, o.OrderId as OrderId, orderdate as OrderDate, ordervalue as OrderValue, tax as Tax, freight as Freight, currency Currency from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId ORDER BY OrderDate, OrderValue desc", sqlConn))
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