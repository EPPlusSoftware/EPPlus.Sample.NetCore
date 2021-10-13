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
using OfficeOpenXml;
using System.Data.SqlClient;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Data.SQLite;

namespace EPPlusSamples.SalesReport
{
    class SalesReportFromDatabase
    {
        /// <summary>
        /// Sample 3 - Creates a workbook and populates using data from a SQLite database
        /// </summary>
        /// <param name="outputDir">The output directory</param>
        /// <param name="templateDir">The location of the sample template</param>
        /// <param name="connectionString">The connection string to the SQLite database</param>
        public static string Run(string connectionString)
        {
            var file = FileOutputUtil.GetFileInfo("08-Salesreport.xlsx");
            using (ExcelPackage xlPackage = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.Add("Sales");
                var namedStyle = xlPackage.Workbook.Styles.CreateNamedStyle("HyperLink");
                namedStyle.Style.Font.UnderLine = true;
                namedStyle.Style.Font.Color.SetColor(Color.Blue);
                const int startRow = 5;
                int row = startRow;
                //Create Headers and format them 
                worksheet.Cells["A1"].Value = "Fiction Inc.";
                using (ExcelRange r = worksheet.Cells["A1:G1"])
                {
                    r.Merge = true;
                    r.Style.Font.SetFromFont(new Font("Britannic Bold", 22, FontStyle.Italic));
                    r.Style.Font.Color.SetColor(Color.White);
                    r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                    r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93));
                }
                worksheet.Cells["A2"].Value = "Sales Report";
                using (ExcelRange r = worksheet.Cells["A2:G2"])
                {
                    r.Merge = true;
                    r.Style.Font.SetFromFont(new Font("Britannic Bold", 18, FontStyle.Italic));
                    r.Style.Font.Color.SetColor(Color.Black);
                    r.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                    r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                }

                worksheet.Cells["A4"].Value = "Company";
                worksheet.Cells["B4"].Value = "Sales Person";
                worksheet.Cells["C4"].Value = "Country";
                worksheet.Cells["D4"].Value = "Order Id";
                worksheet.Cells["E4"].Value = "OrderDate";
                worksheet.Cells["F4"].Value = "Order Value";
                worksheet.Cells["G4"].Value = "Currency";
                worksheet.Cells["A4:G4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["A4:G4"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                worksheet.Cells["A4:G4"].Style.Font.Bold = true;


                //Lets connect to the sample database for some data
                using (var sqlConn = new SQLiteConnection(connectionString))
                {
                    sqlConn.Open();
                    using (var sqlCmd = new SQLiteCommand("select CompanyName, [Name], Email, c.Country, o.OrderId, orderdate, ordervalue, currency from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId ORDER BY 1,2 desc", sqlConn))
                    {
                        using (var sqlReader = sqlCmd.ExecuteReader())
                        {
                            // get the data and fill rows 5 onwards
                            while (sqlReader.Read())
                            {
                                int col = 1;
                                // our query has the columns in the right order, so simply
                                // iterate through the columns
                                for (int i = 0; i < sqlReader.FieldCount; i++)
                                {
                                    // use the email address as a hyperlink for column 1
                                    if (sqlReader.GetName(i) == "email")
                                    {
                                        // insert the email address as a hyperlink for the name
                                        string hyperlink = "mailto:" + sqlReader.GetValue(i).ToString();
                                        worksheet.Cells[row, 2].Hyperlink = new Uri(hyperlink, UriKind.Absolute);
                                    }
                                    else
                                    {
                                        // do not bother filling cell with blank data (also useful if we have a formula in a cell)
                                        if (sqlReader.GetValue(i) != null)
                                            worksheet.Cells[row, col].Value = sqlReader.GetValue(i);
                                        col++;
                                    }
                                }
                                row++;
                            }
                            sqlReader.Close();

                            worksheet.Cells[startRow, 2, row - 1, 2].StyleName = "HyperLink";
                            worksheet.Cells[startRow, 5, row - 1, 5].Style.Numberformat.Format = "yyyy/mm/dd";
                            worksheet.Cells[startRow, 6, row - 1, 6].Style.Numberformat.Format = "[$$-409]#,##0";

                            //Set column width
                            worksheet.Columns[1].Width = 35;
                            worksheet.Columns[2, 3].Width = 28;
                            worksheet.Columns[4].Width = 10;
                            worksheet.Columns[5, 7].Width = 12;
                        }
                    }
                    sqlConn.Close();

                    // lets set the header text 
                    worksheet.HeaderFooter.OddHeader.CenteredText = "Fiction Inc. Sales Report";
                    // add the page number to the footer plus the total number of pages
                    worksheet.HeaderFooter.OddFooter.RightAlignedText =
                        string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                    // add the sheet name to the footer
                    worksheet.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                    // add the file path to the footer
                    worksheet.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FilePath + ExcelHeaderFooter.FileName;
                }
                // we had better add some document properties to the spreadsheet 

                // set some core property values
                xlPackage.Workbook.Properties.Title = "Sales Report";
                xlPackage.Workbook.Properties.Author = "Jan Källman";
                xlPackage.Workbook.Properties.Subject = "Sales Report Samples";
                xlPackage.Workbook.Properties.Keywords = "Office Open XML";
                xlPackage.Workbook.Properties.Category = "Sales Report  Samples";
                xlPackage.Workbook.Properties.Comments = "This sample demonstrates how to create an Excel file from scratch using EPPlus";

                // set some extended property values
                xlPackage.Workbook.Properties.Company = "Fiction Inc.";
                xlPackage.Workbook.Properties.HyperlinkBase = new Uri("https://EPPlusSoftware.com");

                // set some custom property values
                xlPackage.Workbook.Properties.SetCustomPropertyValue("Checked by", "Jan Källman");
                xlPackage.Workbook.Properties.SetCustomPropertyValue("EmployeeID", "1");
                xlPackage.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "EPPlus");

                // save the new spreadsheet
                xlPackage.Save();
            }

            return file.FullName;
        }
    }
}
