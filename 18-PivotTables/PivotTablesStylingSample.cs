﻿/*************************************************************************************************
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
using OfficeOpenXml.Drawing;
using System.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace EPPlusSamples.PivotTables
{
    /// <summary>
    /// This class shows how to use pivottables 
    /// </summary>
    public static class PivotTablesStylingSample
    {
        public static string Run()
        {
            FileInfo templateFile = FileUtil.GetFileInfo("18-PivotTables.xlsx");
            FileInfo newFile = FileUtil.GetFileInfo("18-PivotTables-Styling.xlsx");
            using ( ExcelPackage pck = new ExcelPackage(newFile, templateFile))
            {
                //These two sample shows how to style different parts on the pivot table using pivot areas.
                StylePivotTable1_PerCountry(pck);
                StylePivotTable2_WithDataGrouping(pck);

                //This sample styles the pivot table by combining a named style and use pivot areas. For named styles please also see sample 27
                StylePivotTable3_WithPageFilter(pck);
                //This sample styles the pivot table using pivot areas in various ways and create a custom named slicer style for the slicers.
                StylePivotTable4_WithASlicer(pck);

                //Adds a slicer and do some styling.
                StylePivotTable5_WithACalculatedField(pck);                
                //Sets the pivot table into tabular mode to display the filter boxes on the row fields then styles the button fields
                StylePivotTable6_CaptionFilter(pck);

                //Sets column fields to different background colors.
                StylePivotTable7_WithDataFieldsUsingShowAs(pck);
                
                StylePivotTable8_Sort(pck);
                pck.Save();
            }
            return newFile.FullName;
        }

        private static void StylePivotTable8_Sort(ExcelPackage pck)
        {
            var wsPivot = pck.Workbook.Worksheets["PivotSorting"];

            //Mark the sorted ranges
            var pt1 = wsPivot.PivotTables[0];
            var style1 = pt1.Styles.AddLabel(pt1.RowFields[0]);
            style1.Style.Border.BorderAround(ExcelBorderStyle.DashDotDot,Color.Red);
            
            //Set color red for data cells q
            var pt2 = wsPivot.PivotTables[1];
            var style2 = pt2.Styles.AddAllData();
            style2.GrandColumn = true;
            style2.Style.Font.Color.SetColor(Color.Red);
            
            //Mark sorted column with value "Poland"
            var pt3 = wsPivot.PivotTables[2];
            var style3 = pt3.Styles.AddData(pt3.RowFields[0]);      //Add style for Row field "Name"
            style3.Conditions.DataFields.Add(pt3.DataFields[0]);    //..then add a condition for data field "Order Value" and...
            var conditions = style3.Conditions.Fields.Add(pt3.ColumnFields[0]); //...column field Country with value...
            //pt3.ColumnFields[0].Items.Refresh();
            conditions.Items.AddByValue("Poland");  //..."Poland""..
            style3.Style.Font.Color.SetColor(Color.Green);

        }

        private static void StylePivotTable1_PerCountry(ExcelPackage pck)
        {
            var pivot1 = pck.Workbook.Worksheets["PivotSimple"].PivotTables[0];
            //First add a style that sets the font and color for the entire pivot table.
            var styleWholeTable = pivot1.Styles.AddWholeTable();
            styleWholeTable.Style.Font.Name = "Times New Roman";
            styleWholeTable.Style.Font.Color.SetColor(eThemeSchemeColor.Accent2);

            //Adds new style for all labels in the pivot table. Later added styles will override earlier added styles.
            var styleLabels = pivot1.Styles.AddAllLabels();
            styleLabels.Style.Font.Color.SetColor(eThemeSchemeColor.Accent4);
            styleLabels.Style.Font.Italic = true;

            //This style sets the colors for the labels of the first row field. 
            var styleLabelsForRowField = pivot1.Styles.AddLabel(pivot1.RowFields[0]);
            styleLabelsForRowField.Style.Font.Color.SetColor(eThemeSchemeColor.Text1);

            //This style sets the color and font italic for the grand row of the first row field.
            var styleLabelsForGrandTotal = pivot1.Styles.AddLabel(pivot1.RowFields[0]);
            styleLabelsForGrandTotal.Style.Font.Color.SetColor(Color.Red);
            styleLabelsForGrandTotal.Style.Font.Italic = true;
            styleLabelsForGrandTotal.GrandRow = true;

            //Set the style of the grand total for the data.
            var styleDataForGrandTotal = pivot1.Styles.AddData();
            styleDataForGrandTotal.Style.Font.Color.SetColor(eThemeSchemeColor.Accent6);
            styleDataForGrandTotal.GrandRow = true;
        }
        private static void StylePivotTable2_WithDataGrouping(ExcelPackage pck)
        {
            var pivot2 = pck.Workbook.Worksheets["PivotDateGrp"].PivotTables[0];
            
            //Add a gradient fill for the page field label.
            var stylePagebutton =  pivot2.Styles.AddButtonField(ePivotTableAxis.PageAxis);            
            stylePagebutton.Style.Fill.Style = eDxfFillStyle.GradientFill;
            stylePagebutton.Style.Fill.Gradient.Degree = 90;
            var c1=stylePagebutton.Style.Fill.Gradient.Colors.Add(0);
            c1.Color.SetColor(Color.LightSteelBlue);
            var c2 = stylePagebutton.Style.Fill.Gradient.Colors.Add(1);
            c2.Color.SetColor(Color.DarkSlateBlue);
            stylePagebutton.Style.Font.Color.SetColor(eThemeSchemeColor.Text1);

            //Sets the style for the page filter cell
            var pageStyle = pivot2.Styles.AddLabel(pivot2.PageFields[0]);
            pageStyle.Style.Fill.BackgroundColor.SetColor(Color.DarkGreen);
            stylePagebutton.Style.Font.Color.SetColor(eThemeSchemeColor.Text1);

            //Styles the area to the left of the column axis button field.
            var topLeft = pivot2.Styles.AddTopStart();
            topLeft.Style.Fill.BackgroundColor.SetColor(Color.Green);

            //Set the style for the column axis button field label
            var columnStyle = pivot2.Styles.AddButtonField(ePivotTableAxis.ColumnAxis);
            columnStyle.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            columnStyle.Style.Font.Color.SetColor(eThemeSchemeColor.Text1);

            //Styles the area to the right of the the column axis button field label. 
            var topRight = pivot2.Styles.AddTopEnd();
            topRight.Style.Fill.BackgroundColor.SetColor(Color.Red);




            var rowLableStyleQuarter = pivot2.Styles.AddLabel(pivot2.Fields["Quarters"]);
            rowLableStyleQuarter.Style.Font.Italic = true;

            var rowLableStyleYear = pivot2.Styles.AddLabel(pivot2.Fields["Years"]);
            rowLableStyleYear.Style.Font.Underline = ExcelUnderLineType.Single;

            //Here we style a label for a single row item. We add all the row fields to the pivot area and then add the values we want to style. Note that the value and data type must match the value in the pivot field.
            var labelItem1 = pivot2.Styles.AddLabel(pivot2.Fields["Name"], pivot2.Fields["Years"], pivot2.Fields["Quarters"]);
            labelItem1.Conditions.Fields[0].Items.AddByValue("Christina Parker");
            labelItem1.Conditions.Fields[1].Items.AddByValue(2017D);    //Double here to match the value in the pivot table
            labelItem1.Conditions.Fields[2].Items.AddByValue("Q4");
            labelItem1.Style.Font.Color.SetColor(Color.DarkRed);

            //Here we style a data cell for a single row item. We add all the row fields and the data fields we want to the pivot area and then add the values of the row fields. 
            var dataItem1 = pivot2.Styles.AddData(pivot2.Fields["Name"], pivot2.Fields["Years"], pivot2.Fields["Quarters"]);
            dataItem1.Conditions.Fields[0].Items.AddByValue("Hellen Kuhlman");
            dataItem1.Conditions.Fields[1].Items.AddByValue(2017D);    //Double here to match the value in the pivot table
            dataItem1.Conditions.Fields[2].Items.AddByValue("Q3");
            dataItem1.Conditions.Fields[2].Items.AddByValue("Q4");
            dataItem1.Conditions.DataFields.Add(pivot2.DataFields[0]);  //OrderValue
            dataItem1.Conditions.DataFields.Add(pivot2.DataFields[2]);  //Freight
            dataItem1.Style.Font.Color.SetColor(Color.DarkMagenta);
        }
        private static void StylePivotTable3_WithPageFilter(ExcelPackage pck)
        {
            var pivot3 = pck.Workbook.Worksheets["PivotWithPageField"].PivotTables[0];

            //Create a named pivot table style with Dark28 to start from and make some minor changes.
            var styleName = "CustomPivotTableStyle1";
            var style = pck.Workbook.Styles.CreatePivotTableStyle(styleName, PivotTableStyles.Dark28);
            style.HeaderRow.Style.Font.Italic = true;
            style.TotalRow.Style.Font.Italic = true;
            pivot3.StyleName = styleName;

            //Set the style for the header of the data fields.
            var style1 = pivot3.Styles.AddLabel();
            style1.Conditions.DataFields.Add(pivot3.DataFields[0]);
            style1.Conditions.DataFields.Add(pivot3.DataFields[1]);
            style1.Conditions.DataFields.Add(pivot3.DataFields[2]);
            style1.Style.Font.Underline = ExcelUnderLineType.Single;

            //Here we mark the grand total cell for the last data column.
            var style2 = pivot3.Styles.AddData();
            style2.Conditions.DataFields.Add(pivot3.DataFields[2]);
            style2.GrandRow = true;
            style2.Style.Font.Color.SetColor(Color.Red);

            //Here we set the number format for the OrderDate items for a specific name.
            var style3 = pivot3.Styles.AddData(pivot3.Fields["Name"], pivot3.Fields["OrderDate"]);
            style3.Conditions.Fields[0].Items.AddByValue("Jason Zemlak");
            style3.Conditions.DataFields.Add(pivot3.DataFields[2]);
            style3.Style.NumberFormat.Format="#,##0.00";

            //Here we set the number format of the total cell only.
            var style4 = pivot3.Styles.AddData(pivot3.Fields["Name"]);
            style4.Conditions.Fields[0].Items.AddByValue("Jason Zemlak");
            style4.Conditions.DataFields.Add(pivot3.DataFields[2]);
            style4.Style.NumberFormat.Format = "#,##0.00000";            
            style4.CollapsedLevelsAreSubtotals = true; //Only for the total only. Setting this to false will set the format for the sub items as well
        }

        private static void StylePivotTable4_WithASlicer(ExcelPackage pck)
        {
            //This method connects a slicer to the pivot table. Also see sample 24 for more detailed samples on slicers.
            var wsPivot4 = pck.Workbook.Worksheets["PivotWithSlicer"];
            var pivotTable4= wsPivot4.PivotTables[0];

            //Slicers can also be styled by creating a named style. Here we use the build in Light 5 as a template and changes the font of the slicer.
            //See Sample 27 for more detailed samples.
            var slicer = pivotTable4.Fields["CompanyName"].Slicer;
            var styleName = "CustomSlicerStyle1";
            var style = pck.Workbook.Styles.CreateSlicerStyle(styleName, eSlicerStyle.Light5);
            style.WholeTable.Style.Font.Name = "Stencil";
            slicer.StyleName = styleName;

            var style1 = pivotTable4.Styles.Add();
            style1.GrandRow = true;     //The pivot area will apply to the Grand Row only.
            style1.DataOnly = false;    //DataOnly is true by default, so to apply the style to the entire row we set it to false.
            style1.Style.Font.Size = 18;
        }
        private static void StylePivotTable5_WithACalculatedField(ExcelPackage pck)
        {
            //This method connects a slicer to the pivot table. Also see sample 24 for more detailed samples on slicers.
            var wsPivot5 = pck.Workbook.Worksheets["PivotWithCalculatedField"];
            //Create a new pivot table using the same cache as pivot table 2.
            var pivotTable5 = wsPivot5.PivotTables[0];

            //Sets the entire calculated column fill to solid - Accent 4
            
            //This sets the top-right column. Offset C1 means means offset 3 columns offset within the top end area.
            var style1 = pivotTable5.Styles.AddTopEnd("C1");
            style1.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent4);

            //Sets the fill for the label
            var style2 = pivotTable5.Styles.AddLabel();
            style2.Conditions.DataFields.Add(pivotTable5.DataFields[3]); //Adds a style for the calculated field
            style2.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent4);

            //Sets the fill for the data part of the calculated field.
            var style3 = pivotTable5.Styles.Add();
            style3.Conditions.DataFields.Add(pivotTable5.DataFields[3]); //Add a style for the calculated field
            style3.LabelOnly = false;
            style3.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent4);
        }
        private static void StylePivotTable6_CaptionFilter(ExcelPackage pck)
        {
            var wsPivot6 = pck.Workbook.Worksheets["PivotWithCaptionFilter"];
            var pivotTable6 = wsPivot6.PivotTables[0];

            //Set the pivot table labels in tabular form to get the filter buttons for all row fields.
            pivotTable6.SetCompact(false);

            //Set the style for the buttons.
            var style1 = pivotTable6.Styles.AddButtonField(pivotTable6.RowFields[0]);
            style1.Style.Font.Color.SetColor(eThemeSchemeColor.Accent4);

            var style2 = pivotTable6.Styles.AddButtonField(ePivotTableAxis.RowAxis, 1); //Field with index 1 in the row axis.
            style2.Style.Font.Color.SetColor(eThemeSchemeColor.Accent4);
        }
        private static void StylePivotTable7_WithDataFieldsUsingShowAs(ExcelPackage pck)
        {
            var wsPivot7 = pck.Workbook.Worksheets["PivotWithShowAsFields"];
            var pivotTable7 = wsPivot7.PivotTables[0];

            pivotTable7.PivotTableStyle = PivotTableStyles.Dark18;

            var styleUSD=pivotTable7.Styles.AddData(pivotTable7.Fields["Currency"]);
            styleUSD.Conditions.Fields[0].Items.AddByValue("USD");
            styleUSD.Style.Fill.PatternType = ExcelFillStyle.Solid;            
            styleUSD.Style.Fill.BackgroundColor.Tint = -0.9;

            var styleEUR = pivotTable7.Styles.AddData(pivotTable7.Fields["Currency"]);
            styleEUR.Conditions.Fields[0].Items.AddByValue("EUR");
            styleEUR.Style.Fill.PatternType = ExcelFillStyle.Solid;
            styleEUR.Style.Fill.BackgroundColor.Tint = -0.85;

            var styleSEK = pivotTable7.Styles.AddData(pivotTable7.Fields["Currency"]);
            styleSEK.Conditions.Fields[0].Items.AddByValue("SEK");
            styleSEK.Style.Fill.PatternType = ExcelFillStyle.Solid;
            styleSEK.Style.Fill.BackgroundColor.Tint = -0.80;

            var styleDKK = pivotTable7.Styles.AddData(pivotTable7.Fields["Currency"]);
            styleDKK.Conditions.Fields[0].Items.AddByValue("DKK");
            styleDKK.Style.Fill.PatternType = ExcelFillStyle.Solid;
            styleDKK.Style.Fill.BackgroundColor.Tint = -0.75;

            var styleINR = pivotTable7.Styles.AddData(pivotTable7.Fields["Currency"]);
            styleINR.Conditions.Fields[0].Items.AddByValue("INR");
            styleINR.Style.Fill.PatternType = ExcelFillStyle.Solid;
            styleINR.Style.Fill.BackgroundColor.Tint = -0.70;

            var styleTotal = pivotTable7.Styles.AddData(pivotTable7.Fields["Currency"]);
            styleTotal.GrandRow = true;
            styleTotal.Style.Fill.BackgroundColor.Tint = -1;
        }
    }
}
