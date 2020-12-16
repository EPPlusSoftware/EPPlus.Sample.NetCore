using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Drawing;
using System.Text;

namespace EPPlusSamples
{
    public class FormControlsSample
    {
        public static void Run()
        {
            using (var package = new ExcelPackage())
            {
                //First create the sheet containing the data for the check box and the list box.
                var dataSheet = CreateDataSheet(package);

                //Create the form-sheet and set headers and som basic properties.
                var formSheet = package.Workbook.Worksheets.Add("Form");
                formSheet.Cells["A1"].Value = "Room booking";
                formSheet.Cells["A1"].Style.Font.Size = 18;
                formSheet.Cells["A1"].Style.Font.Bold = true;
                formSheet.Column(1).Width = 30;
                formSheet.Column(2).Width = 60;
                formSheet.View.ShowGridLines = false;
                formSheet.View.ShowHeaders = false;
                formSheet.Cells.Style.Fill.SetBackground(Color.Gray);
                formSheet.DefaultRowHeight = 25;

                //Add texts and format the text fields style
                formSheet.Cells["A3"].Value = "Name";
                formSheet.Cells["A4"].Value = "Gender";
                formSheet.Cells["B3,B5,B11"].Style.Border.BorderAround(ExcelBorderStyle.Dotted);
                formSheet.Cells["B3,B5,B11"].Style.Fill.SetBackground(eThemeSchemeColor.Background1);
                
                //Controls are added via the worksheets drawings collection. 
                //Each type has its typed method returning the specific control class. 
                //Optionally you can use the AddControl method specifying the control type via the eControlType enum
                var dropDown = formSheet.Drawings.AddDropDownControl("DropDown1");
                dropDown.InputRange = dataSheet.Cells["A1:A2"];     //Linkes the range with items
                dropDown.LinkedCell = formSheet.Cells["C4"];        //The cell where the selected index is updated.
                dropDown.SetPosition(3, 1, 1, 0);
                dropDown.SetSize(453, 32);
                
                formSheet.Cells["A5"].Value = "Number of guests";

                //Add a spin button for the number of guests cell
                var spinnButton = formSheet.Drawings.AddSpinButtonControl("SpinButton1");
                spinnButton.SetPosition(4, 0, 2, 1);
                spinnButton.SetSize(32, 35);
                spinnButton.Value = 0;
                spinnButton.Increment = 1;
                spinnButton.MinValue = 0;
                spinnButton.MaxValue = 3;
                spinnButton.LinkedCell = formSheet.Cells["B5"];
                spinnButton.Value = 1;

                //Add a group box and four option boxes to select room type
                var grpBox = formSheet.Drawings.AddGroupBoxControl("GroupBox 1");
                grpBox.Text = "Room types";
                grpBox.SetPosition(5, 8, 1, 1);
                grpBox.SetSize(150, 150);

                var r1 = formSheet.Drawings.AddRadioButtonControl("OptionSingleRoom");
                r1.Text = "Single Room";
                r1.FirstButton = true;
                r1.LinkedCell = formSheet.Cells["C7"];
                r1.SetPosition(5, 15, 1, 5);

                var r2 = formSheet.Drawings.AddRadioButtonControl("OptionDoubleRoom");
                r2.Text = "Double Room";
                r2.LinkedCell = formSheet.Cells["C7"];
                r2.SetPosition(6, 15, 1, 5);
                r2.Checked = true;

                var r3 = formSheet.Drawings.AddRadioButtonControl("OptionSuperiorRoom");
                r3.Text = "Superior";
                r3.LinkedCell = formSheet.Cells["C7"];
                r3.SetPosition(7, 15, 1, 5);

                var r4 = formSheet.Drawings.AddRadioButtonControl("OptionSuite");
                r4.Text = "Suite";
                r4.LinkedCell = formSheet.Cells["C7"];
                r4.SetPosition(8, 15, 1, 5);

                //Group the radio buttons together with the radio buttons, so they act as one unit.
                //You can group drawings via the Group method on one of the drawings...
                var grp = grpBox.Group(r1, r2, r3);     //This will group the groupbox and three of the radio buttons. You would normaly include r4 here as well, but we add it in the next statment to demonstrate how drawings can be grouped.
                //...Or add them to a group drawing returned by the Group method.
                grp.Drawings.Add(r4); //This will add the fourth radio button to the group

                //Add a scroll bar to control the number of nights
                formSheet.Cells["A11"].Value = "Number of nights";
                var scrollBar = formSheet.Drawings.AddScrollBarControl("Scrollbar1");
                scrollBar.Horizontal = true;    //We want a horizontal scrollbar
                scrollBar.SetPosition(10, 1, 2, 1);
                scrollBar.SetSize(200, 30);
                scrollBar.LinkedCell = formSheet.Cells["B11"];
                scrollBar.MinValue = 1;
                scrollBar.MaxValue = 365;
                scrollBar.Increment = 1;
                scrollBar.Page = 7; //How much a page click should increase.
                scrollBar.Value = 1;

                //Add a listbox and connect it to the input range in the data sheet
                formSheet.Cells["A12"].Value = "Requests";
                var listBox = formSheet.Drawings.AddListBoxControl("Listbox1");
                listBox.InputRange = dataSheet.Cells["B1:B3"];
                listBox.LinkedCell = formSheet.Cells["C12"];
                listBox.SetPosition(11, 5, 1, 0);
                listBox.SetSize(200, 100);

                //Last, add a button and connect it to a macro appending the data to a text file.
                var button = formSheet.Drawings.AddButtonControl("ExportButton");
                button.Text = "Make Reservation";
                button.Macro = "ExportButton_Click";
                button.SetPosition(15, 0, 1, 0);
                button.AutomaticSize = true;
                formSheet.Select(formSheet.Cells["B3"]);

                package.Workbook.CreateVBAProject();
                var module = package.Workbook.VbaProject.Modules.AddModule("ControlEvents");
                var code = new StringBuilder();
                code.AppendLine("Sub ExportButton_Click");
                code.AppendLine("Msgbox \"Here you can place the code to handle the form\"");
                code.AppendLine("End Sub");
                module.Code = code.ToString();

                package.SaveAs(FileOutputUtil.GetFileInfo("26-FormControls.xlsm"));
            }
        }

        private static ExcelWorksheet CreateDataSheet(ExcelPackage package)
        {
            var dataSheet = package.Workbook.Worksheets.Add("Data");
            dataSheet.Cells["A1"].Value = "Man";
            dataSheet.Cells["A2"].Value = "Woman";

            dataSheet.Cells["B1"].Value = "Garden view";
            dataSheet.Cells["B2"].Value = "Sea view";
            dataSheet.Cells["B3"].Value = "Parking lot view";

            dataSheet.Hidden = eWorkSheetHidden.Hidden; //We hide the data sheet.

            return dataSheet;
        }
    }
}
