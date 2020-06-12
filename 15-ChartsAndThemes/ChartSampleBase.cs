using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing.Drawing2D;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public abstract class ChartSampleBase
    {
        private class RegionalSales
        {
            public string Region { get; set; }
            public int SoldUnits { get; set; }
            public double TotalSales { get; set; }
            public double Margin { get; set; }
        }
        public static DataTable GetCarDataTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("Car", typeof(string));
            dt.Columns.Add("Acceleration Index", typeof(int));
            dt.Columns.Add("Size Index", typeof(int));
            dt.Columns.Add("Polution Index", typeof(int));
            dt.Columns.Add("Retro Index", typeof(int));
            dt.Rows.Add("Volvo 242", 1, 3, 4, 4);
            dt.Rows.Add("Lamborghini Countach", 5, 1, 5, 4);
            dt.Rows.Add("Tesla Model S", 5, 2, 1, 1);
            dt.Rows.Add("Hummer H1", 2, 5, 5, 2);

            return dt;
        }

        protected static async Task<ExcelRangeBase> LoadFromDatabase(string connectionString, ExcelWorksheet ws)
        {
            ExcelRangeBase range;
            using (var sqlConn = new SQLiteConnection(connectionString))
            {
                sqlConn.Open();
                using (var sqlCmd = new SQLiteCommand("select orderdate as OrderDate, SUM(ordervalue) as OrderValue, SUM(tax) As Tax,SUM(freight) As Freight from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId Where Currency='USD' group by OrderDate ORDER BY OrderDate desc limit 15", sqlConn))
                {
                    using (var sqlReader = sqlCmd.ExecuteReader())
                    {
                        range = await ws.Cells["A1"].LoadFromDataReaderAsync(sqlReader, true);
                        range.Offset(0, 0, 1, range.Columns).Style.Font.Bold = true;
                        range.Offset(0, 0, range.Rows, 1).Style.Numberformat.Format = "yyyy-MM-dd";
                    }
                    //Set the numberformat
                }
            }
            return range;
        }
        protected static async Task<ExcelRangeBase> LoadSalesFromDatabase(string connectionString, ExcelWorksheet ws)
        {
            ExcelRangeBase range;
            using (var sqlConn = new SQLiteConnection(connectionString))
            {
                sqlConn.Open();
                using (var sqlCmd = new SQLiteCommand("select s.continent, s.country, s.city, SUM(OrderValue) As Sales from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId Where Currency='USD' group by s.continent, s.country, s.city ORDER BY s.continent, s.country, s.city", sqlConn))
                {
                    using (var sqlReader = sqlCmd.ExecuteReader())
                    {
                        range = await ws.Cells["A1"].LoadFromDataReaderAsync(sqlReader, true);
                        range.Offset(0, 0, 1, range.Columns).Style.Font.Bold = true;
                        range.Offset(0, 3, range.Rows, 3).Style.Numberformat.Format = "#,##0";
                    }
                    //Set the numberformat
                }
            }
            return range;
        }

        protected static void CreateIceCreamData(ExcelWorksheet ws)
        {
            ws.SetValue("A1", "Icecream Sales-2019");
            ws.SetValue("A2", "Date");
            ws.SetValue("B2", "Sales");
            ws.SetValue("A3", new DateTime(2019, 1, 1));
            ws.SetValue("B3", 2500);
            ws.SetValue("A4", new DateTime(2019, 2, 1));
            ws.SetValue("B4", 3000);
            ws.SetValue("A5", new DateTime(2019, 3, 1));
            ws.SetValue("B5", 2700);
            ws.SetValue("A6", new DateTime(2019, 4, 1));
            ws.SetValue("B6", 4400);
            ws.SetValue("A7", new DateTime(2019, 5, 1));
            ws.SetValue("B7", 6900);
            ws.SetValue("A8", new DateTime(2019, 6, 1));
            ws.SetValue("B8", 11200);
            ws.SetValue("A9", new DateTime(2019, 7, 1));
            ws.SetValue("B9", 13200);
            ws.SetValue("A10", new DateTime(2019, 8, 1));
            ws.SetValue("B10", 12400);
            ws.SetValue("A11", new DateTime(2019, 9, 1));
            ws.SetValue("B11", 8700);
            ws.SetValue("A12", new DateTime(2019, 10, 1));
            ws.SetValue("B12", 4800);
            ws.SetValue("A13", new DateTime(2019, 11, 1));
            ws.SetValue("B13", 2000);
            ws.SetValue("A14", new DateTime(2019, 12, 1));
            ws.SetValue("B14", 2400);
            ws.Cells["A3:A14"].Style.Numberformat.Format = "yyyy-MM";
            ws.Cells["B3:B14"].Style.Numberformat.Format = "#,##0kr";
        }
        protected static ExcelWorksheet LoadBubbleChartData(ExcelPackage package)
        {
            var data = new List<RegionalSales>()
            {
                    new RegionalSales(){ Region = "North", SoldUnits=500, TotalSales=4800, Margin=0.200 },
                    new RegionalSales(){ Region = "Central", SoldUnits=900, TotalSales=7330, Margin=0.333 },
                    new RegionalSales(){ Region = "South", SoldUnits=400, TotalSales=3700, Margin=0.150 },
                    new RegionalSales(){ Region = "East", SoldUnits=350, TotalSales=4400, Margin=0.102 },
                    new RegionalSales(){ Region = "West", SoldUnits=700, TotalSales=6900, Margin=0.218 },
                    new RegionalSales(){ Region = "Stockholm", SoldUnits=1200, TotalSales=8250, Margin=0.350 }
            };
            var wsData = package.Workbook.Worksheets.Add("ChartData");
            wsData.Cells["A1"].LoadFromCollection(data, true, TableStyles.Medium15);
            wsData.Cells["B2:C7"].Style.Numberformat.Format = "#,##0";
            wsData.Cells["D2:D7"].Style.Numberformat.Format = "#,##0.00%";

            var shape = wsData.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.Text = "This worksheet contains the data for the bubble-chartsheet";
            shape.SetPosition(1, 0, 6, 0);
            shape.Effect.SetPresetShadow(ePresetExcelShadowType.OuterBottomLeft);
            return wsData;
        }
    }
}
