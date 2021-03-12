using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using System.Data.SQLite;
using System.Drawing;

namespace EPPlusSamples
{
    /// <summary>
    /// This sample demonstrates how to add custom named styles for 
    /// </summary>
    public static class CustomTableSlicerStyleSample
    {
        public static void Run(string connectionString)
        {
            using (var p = new ExcelPackage())
            {
                CreateTableStyles(p);
                CreatePivotTableStyles(p);
                CreateSlicerStyles(p);

                p.SaveAs(FileOutputUtil.GetFileInfo("27-TableAndSlicerStyles.xlsx"));
            }
        }

        private static void CreateTableStyles(ExcelPackage p)
        {
            var wsTables = p.Workbook.Worksheets.Add("CustomStyledTables");

            //Create a custom table style from scratch and adds a fill gradient fill style 
            var customTableStyle1 = "CustomTableStyle1";
            CreateCustomTableStyleFromScratch(p, customTableStyle1);

            //This samples creates a style with the build in table style Dark11 as template and set the header row and table row font to italic.
            var customTableStyle2 = "CustomTableStyleFromDark11";
            CreateCustomTableStyleFromBuildInTableStyle(p, customTableStyle2);

            //This samples creates a style with the build in table style Medium11 as template and set the header row and table row font to italic.
            var customTableStyle3 = "CustomTableAndPivotTableStyleFromDark11";
            CreateCustomTableAndPivotTableStyleFromBuildInStyle(p, customTableStyle3);


            ExcelTable tbl1 = CreateTable(wsTables, "Table1");
            tbl1.StyleName = customTableStyle1;

            ExcelTable tbl2 = CreateTable(wsTables, "Table2", 9, 1);
            tbl2.StyleName = customTableStyle2;

            ExcelTable tbl3 = CreateTable(wsTables, "Table3", 17, 1);
            tbl3.StyleName = customTableStyle3;

            wsTables.Cells.AutoFitColumns();
        }

        private static void CreatePivotTableStyles(ExcelPackage p)
        {
            var wsPivotTable = p.Workbook.Worksheets.Add("CustomStyledPivotTables");

            //Create a pivot table style from scratch.
            var customPivotTableStyle1 = "CustomPivotTableStyle1";
            CreateCustomPivotTableStyleFromScratch(p, customPivotTableStyle1);

            //This samples creates a style with the build in table style Dark11 as template and set the header row and table row font to italic.
            var customPivotTableStyle2 = "CustomPivotTableStyleFromMedium25";
            CreateCustomPivotTableStyleFromBuildInTableStyle(p, customPivotTableStyle2);

            //Create a pivot table and use the named style we created earlier in this sample for both pivot tables and tables.
            var pt1 = CreatePivotTable(wsPivotTable, "PivotTable1", p.Workbook.Worksheets[0].Tables[0], wsPivotTable.Cells["A3"]);
            pt1.StyleName = "CustomTableAndPivotTableStyleFromDark11";

            var pt2 = CreatePivotTable(wsPivotTable, "PivotTable2", p.Workbook.Worksheets[0].Tables[0], wsPivotTable.Cells["A15"]);
            pt2.StyleName = customPivotTableStyle1;

            var pt3 = CreatePivotTable(wsPivotTable, "PivotTable3", p.Workbook.Worksheets[0].Tables[0], wsPivotTable.Cells["A30"]);
            pt3.StyleName = customPivotTableStyle2;

        }

        private static void CreateSlicerStyles(ExcelPackage p)
        {
            var wsSlicers = p.Workbook.Worksheets.Add("CustomStyledSlicers");
            var tbl=CreateTable(wsSlicers, "TableForSlicer1");

            var slicer1=tbl.Columns[0].AddSlicer();
            slicer1.SetPosition(100, 300);

            //Create a slicer style from scratch.
            var customSlicerStyle1 = "CustomSlicerStyleConsole";
            CreateCustomSlicerStyleFromScratch(p, customSlicerStyle1);
            slicer1.StyleName = customSlicerStyle1;

            var slicer2 = tbl.Columns[1].AddSlicer();
            slicer2.SetPosition(100, 500);

            var customSlicerStyle2 = "CustomSlicerStyleFromStyleDark2";
            CreateCustomSlicerStyleFromBuildInStyle(p, customSlicerStyle2);
            slicer2.StyleName = customSlicerStyle2;
        }

        private static void CreateCustomSlicerStyleFromBuildInStyle(ExcelPackage p, string styleName)
        {
            //Create a custom named slicer style that applies users the build in style Dark4 as template and make some minor modifications.
            var customSlicerStyle = p.Workbook.Styles.CreateSlicerStyle(styleName, eSlicerStyle.Dark4);
            customSlicerStyle.WholeTable.Style.Font.Name = "Broadway";
            customSlicerStyle.HeaderRow.Style.Font.Italic = true;
            customSlicerStyle.HeaderRow.Style.Border.Bottom.Color.SetColor(Color.Red);
        }

        private static void CreateCustomSlicerStyleFromScratch(ExcelPackage p, string styleName)
        {
            //Create a named style that applies to slicers with a console feel to the style.
            var customSlicerStyle = p.Workbook.Styles.CreateSlicerStyle(styleName);

            customSlicerStyle.WholeTable.Style.Font.Name = "Consolas";
            customSlicerStyle.WholeTable.Style.Font.Size = 12;
            customSlicerStyle.WholeTable.Style.Font.Color.SetColor(Color.WhiteSmoke);
            customSlicerStyle.WholeTable.Style.Fill.BackgroundColor.SetColor(Color.Black);

            customSlicerStyle.HeaderRow.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            customSlicerStyle.HeaderRow.Style.Font.Color.SetColor(Color.Black);

            customSlicerStyle.SelectedItemWithData.Style.Fill.BackgroundColor.SetColor(Color.Gray);
            customSlicerStyle.SelectedItemWithData.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.DarkGray);

            customSlicerStyle.SelectedItemWithNoData.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xFF, 64, 64, 64));
            customSlicerStyle.SelectedItemWithNoData.Style.Font.Color.SetColor(Color.DarkGray);
            customSlicerStyle.SelectedItemWithNoData.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.DarkGray);

            customSlicerStyle.UnselectedItemWithData.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.DarkGray);
            customSlicerStyle.UnselectedItemWithNoData.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.DarkGray);
            
            customSlicerStyle.UnselectedItemWithNoData.Style.Font.Color.SetColor(Color.DarkGray);

            customSlicerStyle.HoveredSelectedItemWithData.Style.Fill.BackgroundColor.SetColor(Color.DarkGray);                        
            customSlicerStyle.HoveredSelectedItemWithData.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.WhiteSmoke);            
           
            customSlicerStyle.HoveredSelectedItemWithNoData.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.WhiteSmoke);            
            
            customSlicerStyle.HoveredUnselectedItemWithData.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.WhiteSmoke);
            customSlicerStyle.HoveredUnselectedItemWithNoData.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.WhiteSmoke);
        }

        #region Table Styles
        private static void CreateCustomTableStyleFromScratch(ExcelPackage p, string styleName)
        {
            //Create a named style used to tables only.
            var customTableStyle = p.Workbook.Styles.CreateTableStyle(styleName);

            customTableStyle.WholeTable.Style.Font.Color.SetColor(eThemeSchemeColor.Text2);
            customTableStyle.HeaderRow.Style.Font.Bold = true;
            customTableStyle.HeaderRow.Style.Font.Italic = true;

            customTableStyle.HeaderRow.Style.Fill.Style = eDxfFillStyle.GradientFill;
            customTableStyle.HeaderRow.Style.Fill.Gradient.Degree = 90;

            var c1 = customTableStyle.HeaderRow.Style.Fill.Gradient.Colors.Add(0);
            c1.Color.SetColor(Color.LightGreen);

            var c3 = customTableStyle.HeaderRow.Style.Fill.Gradient.Colors.Add(100);
            c3.Color.SetColor(Color.DarkGreen);
        }
        private static void CreateCustomTableStyleFromBuildInTableStyle(ExcelPackage p, string styleName)
        {
            //Create a new custom table style with the build in style Dark11 as template.
            var customTableStyle = p.Workbook.Styles.CreateTableStyle(styleName, TableStyles.Dark11);

            customTableStyle.HeaderRow.Style.Font.Italic = true;
            customTableStyle.TotalRow.Style.Font.Italic = true;

            //Set the stripe size to 2 rows for both the the row stripes
            customTableStyle.FirstRowStripe.BandSize = 2;
            customTableStyle.FirstRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

            customTableStyle.SecondRowStripe.BandSize = 2;
            customTableStyle.SecondRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);
        }
        private static void CreateCustomTableAndPivotTableStyleFromBuildInStyle(ExcelPackage p, string customTableStyle3)
        {
            //Create a named style that can be used both for tables and pivot tables. 
            //We create this style from one of the build in pivot table styles - Medium13, but table styles can also be used as a parameter for this method
            var customTableStyle = p.Workbook.Styles.CreateTableAndPivotTableStyle(customTableStyle3, PivotTableStyles.Medium13);

            //Set the header row and total row border to dotted.
            customTableStyle.HeaderRow.Style.Border.Bottom.Style = ExcelBorderStyle.Dotted;
            customTableStyle.HeaderRow.Style.Border.Bottom.Color.SetColor(Color.Gray);

            customTableStyle.TotalRow.Style.Border.Top.Style = ExcelBorderStyle.Dotted;
            customTableStyle.TotalRow.Style.Border.Top.Color.SetColor(Color.Gray);
        }
        /// <summary>
        /// Creates a table with random data used for this sample
        /// </summary>
        /// <param name="wsTables">The worksheet </param>
        /// <param name="tableName">The name of the table</param>
        /// <param name="rowStart">Start row of the table</param>
        /// <param name="colStart">Start column of the table</param>
        /// <returns></returns>
        private static ExcelTable CreateTable(ExcelWorksheet wsTables, string tableName, int rowStart = 1, int colStart = 1)
        {
            wsTables.Cells[rowStart, colStart].Value = "Column1";
            wsTables.Cells[rowStart, colStart + 1].Value = "Column2";
            wsTables.Cells[rowStart, colStart + 2].Value = "Column3";
            wsTables.Cells[rowStart + 1, colStart].Value = 1;
            wsTables.Cells[rowStart + 1, colStart + 1].Value = 2;
            wsTables.Cells[rowStart + 1, colStart + 2].Value = "Type 1";

            wsTables.Cells[rowStart + 2, colStart].Value = 2;
            wsTables.Cells[rowStart + 2, colStart + 1].Value = 4;
            wsTables.Cells[rowStart + 2, colStart + 2].Value = "Type 2";

            wsTables.Cells[rowStart + 3, colStart].Value = 3;
            wsTables.Cells[rowStart + 3, colStart + 1].Value = 7;
            wsTables.Cells[rowStart + 3, colStart + 2].Value = "Type 1";

            wsTables.Cells[rowStart + 4, colStart].Value = 4;
            wsTables.Cells[rowStart + 4, colStart + 1].Value = 20;
            wsTables.Cells[rowStart + 4, colStart + 2].Value = "Type 3";

            wsTables.Cells[rowStart + 5, colStart].Value = 5;
            wsTables.Cells[rowStart + 5, colStart + 1].Value = 43;
            wsTables.Cells[rowStart + 5, colStart + 2].Value = "Type 3";

            var tbl = wsTables.Tables.Add(wsTables.Cells[rowStart, colStart, rowStart + 5, colStart + 2], tableName);
            tbl.Columns[0].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[1].TotalsRowFunction = RowFunctions.Sum;
            tbl.ShowTotal = true;
            return tbl;
        }
        #endregion
        #region Pivot Table Styles

        private static void CreateCustomPivotTableStyleFromScratch(ExcelPackage p, string styleName)
        {
            //Create a named style that applies only to pivot tables.
            var customPivotTableStyle = p.Workbook.Styles.CreatePivotTableStyle(styleName);

            customPivotTableStyle.WholeTable.Style.Font.Color.SetColor(ExcelIndexedColor.Indexed22);
            customPivotTableStyle.PageFieldLabels.Style.Font.Color.SetColor(Color.Red);
            customPivotTableStyle.PageFieldValues.Style.Font.Color.SetColor(eThemeSchemeColor.Accent4);

            customPivotTableStyle.HeaderRow.Style.Font.Color.SetColor(Color.DarkGray);
            customPivotTableStyle.HeaderRow.Style.Fill.Style = eDxfFillStyle.GradientFill;
            customPivotTableStyle.HeaderRow.Style.Fill.Gradient.Degree = 180;

            var c1 = customPivotTableStyle.HeaderRow.Style.Fill.Gradient.Colors.Add(0);
            c1.Color.SetColor(Color.LightBlue);

            var c3 = customPivotTableStyle.HeaderRow.Style.Fill.Gradient.Colors.Add(100);
            c3.Color.SetColor(Color.DarkCyan);

        }
        private static void CreateCustomPivotTableStyleFromBuildInTableStyle(ExcelPackage p, string styleName)
        {
            //Create a new custom pivot table style with the build in style Medium as template.
            var customPivotTableStyle = p.Workbook.Styles.CreatePivotTableStyle(styleName, PivotTableStyles.Medium25);

            //Alter the font color of the entire table to theme color Text 2
            customPivotTableStyle.WholeTable.Style.Font.Color.SetColor(eThemeSchemeColor.Text2);

            customPivotTableStyle.HeaderRow.Style.Font.Italic = true;
            customPivotTableStyle.TotalRow.Style.Font.Italic = true;

            customPivotTableStyle.FirstColumnSubheading.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
            customPivotTableStyle.FirstColumnStripe.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            customPivotTableStyle.FirstColumnStripe.Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
        }

        private static ExcelPivotTable CreatePivotTable(ExcelWorksheet wsPivotTables, string pivotTableName, ExcelTable tableSource, ExcelRange range)
        {
            var pt = wsPivotTables.PivotTables.Add(range, tableSource, pivotTableName);
            pt.RowFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[1]);
            pt.PageFields.Add(pt.Fields[2]);
            return pt;
        }
        #endregion
    }
}
