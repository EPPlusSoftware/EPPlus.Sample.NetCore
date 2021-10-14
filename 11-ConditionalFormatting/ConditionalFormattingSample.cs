/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  2012-04-03         Eyal Seagull                 Added
  2020-01-26         EPPlus Software AB           Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Drawing;

namespace EPPlusSamples
{
  class ConditionalFormatting
    {
    /// <summary>
    /// Sample 14 - Conditional formatting example
    /// </summary>
    public static string Run()
    {
      FileInfo newFile = FileUtil.GetCleanFileInfo("11-ConditionalFormatting.xlsx");
      using (ExcelPackage package = new ExcelPackage(newFile))
      {
        // add a new worksheet to the empty workbook
        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Conditional Formatting");

        // Create 4 columns of samples data
        for (int col = 1; col < 10; col++)
        {
          // Add the headers
          worksheet.Cells[1, col].Value = "Sample " + col;

          for (int row = 2; row < 21; row++)
          {
            // Add some items...
            worksheet.Cells[row, col].Value = row;
          }
        }

        // -------------------------------------------------------------------
        // TwoColorScale Conditional Formatting example
        // -------------------------------------------------------------------
        ExcelAddress cfAddress1 = new ExcelAddress("A2:A10");
        var cfRule1 = worksheet.ConditionalFormatting.AddTwoColorScale(cfAddress1);

        // Now, lets change some properties:
        cfRule1.LowValue.Type = eExcelConditionalFormattingValueObjectType.Num;
        cfRule1.LowValue.Value = 4;
        cfRule1.LowValue.Color = Color.FromArgb(0xFF, 0xFF, 0xEB, 0x84);
        cfRule1.HighValue.Type = eExcelConditionalFormattingValueObjectType.Formula;
        cfRule1.HighValue.Formula = "IF($G$1=\"A</x:&'cfRule>\",1,5)";
        cfRule1.StopIfTrue = true;
        cfRule1.Style.Font.Bold = true;

        // But others you can't (readonly)
        // cfRule1.Type = eExcelConditionalFormattingRuleType.ThreeColorScale;

        // -------------------------------------------------------------------
        // ThreeColorScale Conditional Formatting example
        // -------------------------------------------------------------------
        ExcelAddress cfAddress2 = new ExcelAddress(2, 2, 10, 2);  //="B2:B10"
        var cfRule2 = worksheet.ConditionalFormatting.AddThreeColorScale(cfAddress2);

        // Changing some properties again
        cfRule2.Priority = 1;
        cfRule2.MiddleValue.Type = eExcelConditionalFormattingValueObjectType.Percentile;
        cfRule2.MiddleValue.Value = 30;
        cfRule2.StopIfTrue = true;

        // You can access a rule by its Priority
        var cfRule2Priority = cfRule2.Priority;
        var cfRule2_1 = worksheet.ConditionalFormatting.RulesByPriority(cfRule2Priority);

        // And you can even change the rule's Address
        cfRule2_1.Address = new ExcelAddress("Z1:Z3");

        // -------------------------------------------------------------------
        // Adding another ThreeColorScale in a different way (observe that we are
        // pointing to the same range as the first rule we entered. Excel allows it to
        // happen and group the rules in one <conditionalFormatting> node)
        // -------------------------------------------------------------------
        var cfRule3 = worksheet.Cells[cfAddress1.Address].ConditionalFormatting.AddThreeColorScale();
        cfRule3.LowValue.Color = Color.LemonChiffon;
        
        // -------------------------------------------------------------------
        // Change the rules priorities to change their execution order
        // -------------------------------------------------------------------
        cfRule3.Priority = 1;
        cfRule1.Priority = 2;
        cfRule2.Priority = 3;

        // -------------------------------------------------------------------
        // Create an Above Average rule
        // -------------------------------------------------------------------
        var cfRule4 = worksheet.ConditionalFormatting.AddAboveAverage(
          new ExcelAddress("B11:B20"));
        cfRule4.Style.Font.Bold = true;
        cfRule4.Style.Font.Color.Color = Color.Red;
        cfRule4.Style.Font.Strike = true;

        // -------------------------------------------------------------------
        // Create an Above Or Equal Average rule
        // -------------------------------------------------------------------
        var cfRule5 = worksheet.ConditionalFormatting.AddAboveOrEqualAverage(
          new ExcelAddress("B11:B20"));

        // -------------------------------------------------------------------
        // Create a Below Average rule
        // -------------------------------------------------------------------
        var cfRule6 = worksheet.ConditionalFormatting.AddBelowAverage(
          new ExcelAddress("B11:B20"));

        // -------------------------------------------------------------------
        // Create a Below Or Equal Average rule
        // -------------------------------------------------------------------
        var cfRule7 = worksheet.ConditionalFormatting.AddBelowOrEqualAverage(
          new ExcelAddress("B11:B20"));

        // -------------------------------------------------------------------
        // Create a Above StdDev rule
        // -------------------------------------------------------------------
        var cfRule8 = worksheet.ConditionalFormatting.AddAboveStdDev(
          new ExcelAddress("B11:B20"));
          cfRule8.StdDev = 0;

        // -------------------------------------------------------------------
        // Create a Below StdDev rule
        // -------------------------------------------------------------------
        var cfRule9 = worksheet.ConditionalFormatting.AddBelowStdDev(
          new ExcelAddress("B11:B20"));

        cfRule9.StdDev = 2;

        // -------------------------------------------------------------------
        // Create a Bottom rule
        // -------------------------------------------------------------------
        var cfRule10 = worksheet.ConditionalFormatting.AddBottom(
          new ExcelAddress("B11:B20"));

        cfRule10.Rank = 4;

        // -------------------------------------------------------------------
        // Create a Bottom Percent rule
        // -------------------------------------------------------------------
        var cfRule11 = worksheet.ConditionalFormatting.AddBottomPercent(
          new ExcelAddress("B11:B20"));

        cfRule11.Rank = 15;

        // -------------------------------------------------------------------
        // Create a Top rule
        // -------------------------------------------------------------------
        var cfRule12 = worksheet.ConditionalFormatting.AddTop(
          new ExcelAddress("B11:B20"));

        // -------------------------------------------------------------------
        // Create a Top Percent rule
        // -------------------------------------------------------------------
        var cfRule13 = worksheet.ConditionalFormatting.AddTopPercent(
          new ExcelAddress("B11:B20"));
        
        cfRule13.Style.Border.Left.Style = ExcelBorderStyle.Thin;
        cfRule13.Style.Border.Left.Color.Theme = eThemeSchemeColor.Text2;
        cfRule13.Style.Border.Bottom.Style = ExcelBorderStyle.DashDot;
        cfRule13.Style.Border.Bottom.Color.SetColor(ExcelIndexedColor.Indexed8);
        cfRule13.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        cfRule13.Style.Border.Right.Color.Color=Color.Blue;
        cfRule13.Style.Border.Top.Style = ExcelBorderStyle.Hair;
        cfRule13.Style.Border.Top.Color.Auto=true;

        // -------------------------------------------------------------------
        // Create a Last 7 Days rule
        // -------------------------------------------------------------------
        ExcelAddress timePeriodAddress = new ExcelAddress("D21:G34 C11:C20");
        var cfRule14 = worksheet.ConditionalFormatting.AddLast7Days(
          timePeriodAddress);

        cfRule14.Style.Fill.PatternType = ExcelFillStyle.LightTrellis;
        cfRule14.Style.Fill.PatternColor.Color = Color.BurlyWood;
        cfRule14.Style.Fill.BackgroundColor.Color = Color.LightCyan;

        // -------------------------------------------------------------------
        // Create a Last Month rule
        // -------------------------------------------------------------------
        var cfRule15 = worksheet.ConditionalFormatting.AddLastMonth(
          timePeriodAddress);

        cfRule15.Style.NumberFormat.Format = "YYYY";
        // -------------------------------------------------------------------
        // Create a Last Week rule
        // -------------------------------------------------------------------
        var cfRule16 = worksheet.ConditionalFormatting.AddLastWeek(
          timePeriodAddress);
        cfRule16.Style.NumberFormat.Format = "YYYY";

        // -------------------------------------------------------------------
        // Create a Next Month rule
        // -------------------------------------------------------------------
        var cfRule17 = worksheet.ConditionalFormatting.AddNextMonth(
          timePeriodAddress);

        // -------------------------------------------------------------------
        // Create a Next Week rule
        // -------------------------------------------------------------------
        var cfRule18 = worksheet.ConditionalFormatting.AddNextWeek(
          timePeriodAddress);

        // -------------------------------------------------------------------
        // Create a This Month rule
        // -------------------------------------------------------------------
        var cfRule19 = worksheet.ConditionalFormatting.AddThisMonth(
          timePeriodAddress);

        // -------------------------------------------------------------------
        // Create a This Week rule
        // -------------------------------------------------------------------
        var cfRule20 = worksheet.ConditionalFormatting.AddThisWeek(
          timePeriodAddress);

        // -------------------------------------------------------------------
        // Create a Today rule
        // -------------------------------------------------------------------
        var cfRule21 = worksheet.ConditionalFormatting.AddToday(
          timePeriodAddress);

        // -------------------------------------------------------------------
        // Create a Tomorrow rule
        // -------------------------------------------------------------------
        var cfRule22 = worksheet.ConditionalFormatting.AddTomorrow(
          timePeriodAddress);

        // -------------------------------------------------------------------
        // Create a Yesterday rule
        // -------------------------------------------------------------------
        var cfRule23 = worksheet.ConditionalFormatting.AddYesterday(
          timePeriodAddress);

        // -------------------------------------------------------------------
        // Create a BeginsWith rule
        // -------------------------------------------------------------------
        ExcelAddress cellIsAddress = new ExcelAddress("E11:E20");
        var cfRule24 = worksheet.ConditionalFormatting.AddBeginsWith(
          cellIsAddress);

        cfRule24.Text = "SearchMe";

        // -------------------------------------------------------------------
        // Create a Between rule
        // -------------------------------------------------------------------
        var cfRule25 = worksheet.ConditionalFormatting.AddBetween(
          cellIsAddress);

        cfRule25.Formula = "IF(E11>5,10,20)";
        cfRule25.Formula2 = "IF(E11>5,30,50)";

        // -------------------------------------------------------------------
        // Create a ContainsBlanks rule
        // -------------------------------------------------------------------
        var cfRule26 = worksheet.ConditionalFormatting.AddContainsBlanks(
          cellIsAddress);

        // -------------------------------------------------------------------
        // Create a ContainsErrors rule
        // -------------------------------------------------------------------
        var cfRule27 = worksheet.ConditionalFormatting.AddContainsErrors(
          cellIsAddress);

        // -------------------------------------------------------------------
        // Create a ContainsText rule
        // -------------------------------------------------------------------
        var cfRule28 = worksheet.ConditionalFormatting.AddContainsText(
          cellIsAddress);

        cfRule28.Text = "Me";

        // -------------------------------------------------------------------
        // Create a DuplicateValues rule
        // -------------------------------------------------------------------
        var cfRule29 = worksheet.ConditionalFormatting.AddDuplicateValues(
          cellIsAddress);

        // -------------------------------------------------------------------
        // Create an EndsWith rule
        // -------------------------------------------------------------------
        var cfRule30 = worksheet.ConditionalFormatting.AddEndsWith(
          cellIsAddress);

        cfRule30.Text = "EndText";

        // -------------------------------------------------------------------
        // Create an Equal rule
        // -------------------------------------------------------------------
        var cfRule31 = worksheet.ConditionalFormatting.AddEqual(
          cellIsAddress);

        cfRule31.Formula = "6";

        // -------------------------------------------------------------------
        // Create an Expression rule
        // -------------------------------------------------------------------
        var cfRule32 = worksheet.ConditionalFormatting.AddExpression(
          cellIsAddress);

        cfRule32.Formula = "E11=E12";

        // -------------------------------------------------------------------
        // Create a GreaterThan rule
        // -------------------------------------------------------------------
        var cfRule33 = worksheet.ConditionalFormatting.AddGreaterThan(
          cellIsAddress);

        cfRule33.Formula = "SE(E11<10,10,65)";

        // -------------------------------------------------------------------
        // Create a GreaterThanOrEqual rule
        // -------------------------------------------------------------------
        var cfRule34 = worksheet.ConditionalFormatting.AddGreaterThanOrEqual(
          cellIsAddress);

        cfRule34.Formula = "35";

        // -------------------------------------------------------------------
        // Create a LessThan rule
        // -------------------------------------------------------------------
        var cfRule35 = worksheet.ConditionalFormatting.AddLessThan(
          cellIsAddress);

        cfRule35.Formula = "36";

        // -------------------------------------------------------------------
        // Create a LessThanOrEqual rule
        // -------------------------------------------------------------------
        var cfRule36 = worksheet.ConditionalFormatting.AddLessThanOrEqual(
          cellIsAddress);

        cfRule36.Formula = "37";

        // -------------------------------------------------------------------
        // Create a NotBetween rule
        // -------------------------------------------------------------------
        var cfRule37 = worksheet.ConditionalFormatting.AddNotBetween(
          cellIsAddress);

        cfRule37.Formula = "333";
        cfRule37.Formula2 = "999";

        // -------------------------------------------------------------------
        // Create a NotContainsBlanks rule
        // -------------------------------------------------------------------
        var cfRule38 = worksheet.ConditionalFormatting.AddNotContainsBlanks(
          cellIsAddress);

        // -------------------------------------------------------------------
        // Create a NotContainsErrors rule
        // -------------------------------------------------------------------
        var cfRule39 = worksheet.ConditionalFormatting.AddNotContainsErrors(
          cellIsAddress);

        // -------------------------------------------------------------------
        // Create a NotContainsText rule
        // -------------------------------------------------------------------
        var cfRule40 = worksheet.ConditionalFormatting.AddNotContainsText(
          cellIsAddress);

        cfRule40.Text = "NotMe";

        // -------------------------------------------------------------------
        // Create an NotEqual rule
        // -------------------------------------------------------------------
        var cfRule41 = worksheet.ConditionalFormatting.AddNotEqual(
          cellIsAddress);

        cfRule41.Formula = "14";

        ExcelAddress cfAddress43 = new ExcelAddress("G2:G10");
        var cfRule42 = worksheet.ConditionalFormatting.AddThreeIconSet(cfAddress43, eExcelconditionalFormatting3IconsSetType.TrafficLights1);

        ExcelAddress cfAddress44 = new ExcelAddress("H2:H10");
        var cfRule43 = worksheet.ConditionalFormatting.AddDatabar(cfAddress44, Color.DarkBlue);
        
          // -----------------------------------------------------------
        // Removing Conditional Formatting rules
        // -----------------------------------------------------------
        // Remove one Rule by its object
        //worksheet.ConditionalFormatting.Remove(cfRule1);

        // Remove one Rule by index
        //worksheet.ConditionalFormatting.RemoveAt(1);

        // Remove one Rule by its Priority
        //worksheet.ConditionalFormatting.RemoveByPriority(2);

        // Remove all the Rules
        //worksheet.ConditionalFormatting.RemoveAll();

        // set some document properties
        package.Workbook.Properties.Title = "Conditional Formatting";
        package.Workbook.Properties.Author = "Eyal Seagull";
        package.Workbook.Properties.Comments = "This sample demonstrates how to add Conditional Formatting to an Excel 2007 worksheet using EPPlus";

        // set some custom property values
        package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Eyal Seagull");
        package.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "EPPlus");

        //Getting a rule from the collection as a typed rule
        if(worksheet.ConditionalFormatting[41].Type==eExcelConditionalFormattingRuleType.ThreeIconSet)
        {
            var iconRule = worksheet.ConditionalFormatting[41].As.ThreeIconSet; //Type cast the rule as an iconset rule.    
            //Do something with the iconRule...
        }
        if (worksheet.ConditionalFormatting[42].Type == eExcelConditionalFormattingRuleType.DataBar)
        {

            var dataBarRule = worksheet.ConditionalFormatting[42].As.DataBar; //Type cast the rule as an iconset rule.
            //Do something with the databarRule...
        }
        // save our new workbook and we are done!
        package.Save();
      }

      return newFile.FullName;
    }
  }
}
