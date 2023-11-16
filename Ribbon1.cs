using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace VSTO_Addins
{


    public partial class Ribbon1
    {

        private Excel.Application excelApp;
        private Excel.Workbook workbook;
        private Excel.Worksheet worksheet;





        // Public Function ConvertImage(ByVal image As Image) As stdole.IPictureDisp
        // Return DirectCast(AxHost.GetIPictureDispFromPicture(image), stdole.IPictureDisp)
        // End Function


        private object SplitText(object Source, object Pattern, object Consecutive, object KeepSeparator, object Before)
        {
            object SplitTextRet = default;

            var SplitValues = new string[1];
            int Index = -1;
            int Start = 1;

            for (int i = 1, loopTo = Strings.Len(Pattern); i <= loopTo; i++)
            {
                if (Strings.Mid(Conversions.ToString(Pattern), i, 1) != "*")
                {
                    int SeparatorLength = 1;
                    while (Strings.Mid(Conversions.ToString(Pattern), i + SeparatorLength, 1) != "*")
                        SeparatorLength = SeparatorLength + 1;
                    string separator = Strings.Mid(Conversions.ToString(Pattern), i, SeparatorLength);
                    int Ending = Strings.InStr(Conversions.ToString(Source), separator);
                    Interaction.MsgBox(Ending);
                    Index = Index + 1;
                    Array.Resize(ref SplitValues, Index + 1);
                    SplitValues[Index] = Strings.Mid(Conversions.ToString(Source), Start, Ending - Start);
                    Start = Ending + Strings.Len(separator);
                }
            }

            SplitTextRet = SplitValues;
            return SplitTextRet;

        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            // Dim form As New Form1
            // form.Show()

            var MyForm1 = new Form1();
            if (GlobalModule.form_flag == false)
            {

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                MyForm1.OpenSheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selection = (Range)excelApp.Selection;

                MyForm1.TextBox1.Text = selection.get_Address();
                MyForm1.ComboBox1.SelectedIndex = -1;
                MyForm1.ComboBox1.Text = "SOFTEKO";

                MyForm1.Show();
                GlobalModule.form_flag = true;
            }

        }

        private void Button2_Click(object sender, RibbonControlEventArgs e)
        {
            // Dim form As New Form2
            // form.Show()
        }

        private void Button3_Click(object sender, RibbonControlEventArgs e)
        {
            // Dim form As New Form3
            // form.Show()
            if (GlobalModule.form_flag == false)   // For avoiding multiple occurrence
            {
                var MyForm3 = new Form3();

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                MyForm3.excelApp = excelApp;
                MyForm3.workbook = workbook;
                MyForm3.worksheet = worksheet;
                MyForm3.workbook2 = workbook;
                MyForm3.worksheet2 = worksheet;
                MyForm3.OpenSheet = worksheet;

                MyForm3.FocusedTextBox = 0;
                MyForm3.Form4Open = 0;
                MyForm3.Workbook2Opened = false;

                Range selection = (Range)excelApp.Selection;

                MyForm3.TextBox1.Text = selection.get_Address();
                MyForm3.ComboBox1.SelectedIndex = -1;
                MyForm3.ComboBox1.Text = "SOFTEKO";

                MyForm3.Show();
                GlobalModule.form_flag = true;
            }

        }

        private void Button5_Click(object sender, RibbonControlEventArgs e)
        {
            // Dim form As New Form8
            // form.Show()
            if (GlobalModule.form_flag == false)
            {
                var MyForm8 = new Form8();

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selection = (Range)excelApp.Selection;

                MyForm8.TextBox1.Text = selection.get_Address();
                MyForm8.ComboBox1.SelectedIndex = -1;
                MyForm8.ComboBox1.Text = "SOFTEKO";
                MyForm8.Show();
                GlobalModule.form_flag = true;
            }

        }

        private void Button6_Click(object sender, RibbonControlEventArgs e)
        {
            // Dim form As New Form10

            // form.Show()
            if (GlobalModule.form_flag == false)
            {
                var MyForm10 = new Form10();

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selection = (Range)excelApp.Selection;

                MyForm10.TextBox1.Text = selection.get_Address();
                MyForm10.ComboBox1.SelectedIndex = -1;
                MyForm10.ComboBox1.Text = "SOFTEKO";
                MyForm10.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button7_Click(object sender, RibbonControlEventArgs e)
        {
            // Dim form As New Form7

            // form.Show()
            if (GlobalModule.form_flag == false)
            {
                var MyForm7 = new Form7();

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selection = (Range)excelApp.Selection;

                MyForm7.TextBox1.Text = selection.get_Address();
                MyForm7.Show();
                GlobalModule.form_flag = true;

            }
        }

        private void Button8_Click(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var form = new Form11SwapRanges();

                form.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button4_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void Button11_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void Button12_Click(object sender, RibbonControlEventArgs e)
        {
            var form = new Form14SpecifyScrollArea();
            form.Show();
        }

        private void Button13_Click(object sender, RibbonControlEventArgs e)
        {
            // MsgBox(form_flag)
            if (GlobalModule.form_flag == false)
            {
                var form = new Form15CompareCells();
                form.Show();
                GlobalModule.form_flag = true;
            }

        }

        private void Button14_Click(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var form = new Form16PasteintoVisibleRange();
                form.Show();
                GlobalModule.form_flag = true;
            }

        }

        private void Button15_Click(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var form = new Form17DivideNames();
                form.Show();
                GlobalModule.form_flag = true;
            }

        }

        private void Button16_Click_1(object sender, RibbonControlEventArgs e)
        {
            // Dim form As New Form18_CombineRanges
            // form.Show()
        }

        private void Button19_Click(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var form = new Form21FillEmtyCells();
                form.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button20_Click(object sender, RibbonControlEventArgs e)
        {
            // If form_flag = False
            // Dim form As New Form22_Merge_Duplicate_Rows
            // form.Show()
        }

        private void Button21_Click(object sender, RibbonControlEventArgs e)
        {
            // Dim form As New Form23_Merge_Duplicate_Columns
            // form.Show()
        }

        private void Button22_Click(object sender, RibbonControlEventArgs e)
        {
            // Dim form As New Form24_Split_Cells
            // form.Show()
        }

        private void Button23_Click(object sender, RibbonControlEventArgs e)
        {
            // Dim form As New Form25_Split_Range
            // form.Show()
        }

        private void Button24_Click(object sender, RibbonControlEventArgs e)
        {
            // Dim form As New Form26_split_text_bycharacters
            // form.Show()
        }

        private void Button25_Click(object sender, RibbonControlEventArgs e)
        {
            // Dim form As New Form27_Split_text_bystrings
            // form.Show()
        }

        private void Button26_Click(object sender, RibbonControlEventArgs e)
        {
            // Dim form As New Form28_Split_text_bypattern
            // form.Show()
        }

        private void Button27_Click(object sender, RibbonControlEventArgs e)
        {
            // Dim form As New Form29_Simple_Drop_down_List
            // form.Show()
        }

        private void Button28_Click(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var form = new Form30_Create_Dynamic_Drop_down_List();

                form.Show();
                GlobalModule.form_flag = true;
            }
        }

        /// <summary>
    /// Unhides all  hidden rows and columns from the entire sheet.
    /// </summary>

        private void Button31_Click(object sender, RibbonControlEventArgs e)
        {

            try
            {

                Excel.Application excelApp;
                Excel.Workbook workbook;
                Excel.Worksheet worksheet;


                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                // unhide all hidden rows and columns
                worksheet.Rows.Hidden = false;
                worksheet.Columns.Hidden = false;
            }


            catch (Exception ex)
            {

            }


        }

        /// <summary>
    /// Unhides any hidden rows and columns from a Selected Range.
    /// </summary>

        private void Button32_Click(object sender, RibbonControlEventArgs e)
        {

            try
            {

                Excel.Application excelApp;
                Excel.Workbook workbook;
                Excel.Worksheet worksheet;
                Range selectedRange;

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                selectedRange = (Range)excelApp.Selection;

                // Define varibales to store row and column numbers of the selected range
                int rowNum, colNum;

                // takes all the ranges selected by user into an array named arrRng
                string[] arrRng = Strings.Split(selectedRange.get_Address(), ",");

                // loops through each range selected by user, which is stored in arrRng array
                for (int p = 0, loopTo = Information.UBound(arrRng); p <= loopTo; p++)
                {

                    selectedRange = worksheet.get_Range(arrRng[p]);

                    // Loop through each row of the selected range
                    // Check if the entire row is hidden
                    // If the row is hidden, unhide it
                    var loopTo1 = selectedRange.Rows.Count;
                    for (rowNum = 1; rowNum <= loopTo1; rowNum++)
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.get_Range(selectedRange.Cells[rowNum, 1], selectedRange.Cells[rowNum, 3]).EntireRow.Hidden, true, false)))
                        {
                            worksheet.get_Range(selectedRange.Cells[rowNum, 1], selectedRange.Cells[rowNum, 3]).EntireRow.Hidden = (object)false;
                        }
                    }

                    // Loop through each column in the selected range
                    // Check if the entire column is hidden
                    // If the column is hidden, unhide it
                    var loopTo2 = selectedRange.Columns.Count;
                    for (colNum = 1; colNum <= loopTo2; colNum++)
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.get_Range(selectedRange.Cells[1, colNum], selectedRange.Cells[3, colNum]).EntireColumn.Hidden, true, false)))
                        {
                            worksheet.get_Range(selectedRange.Cells[1, colNum], selectedRange.Cells[3, colNum]).EntireColumn.Hidden = (object)false;
                        }
                    }
                }
            }

            catch (Exception ex)
            {

            }


        }

        /// <summary>
    /// Removes blank columns from selected range
    /// </summary>
        private void Button33_Click(object sender, RibbonControlEventArgs e)
        {

            try
            {

                Excel.Application excelApp;
                Excel.Workbook workbook;
                Excel.Worksheet worksheet;
                Range selectedRng;
                int blankColCount = 0;
                string flag = "Empty";
                string ValueFlag = "Empty";


                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                selectedRng = (Range)excelApp.Selection;

                // "rngCount" variable indicates the number of ranges in the users' selection.
                // 0 means a single continuous selection
                // > 0 means user selectd multiple disjoint ranges
                int rngCount;
                rngCount = 0;

                foreach (char c in selectedRng.get_Address())
                {

                    if (Conversions.ToString(c) == ",")
                    {
                        rngCount = rngCount + 1;
                    }

                }

                // if user select a single continuous range "rngCount" will be 0
                if (rngCount == 0)
                {

                    // if user select only a single cell the following warning will pop up and exit from the code 
                    if (selectedRng.Rows.Count == 1 & selectedRng.Columns.Count == 1)
                    {
                        Interaction.MsgBox("This Add-in doesn't work for single cell. Please Select a Range and try again!", MsgBoxStyle.Exclamation, "Warning");
                        return;
                    }


                    // loops through the entire selection and if the entire selection is empty or not
                    // if the entire selection is blank then valueFlag will become "NotEmpty"  
                    for (int i = 1, loopTo = selectedRng.Rows.Count; i <= loopTo; i++)
                    {
                        for (int j = 1, loopTo1 = selectedRng.Columns.Count; j <= loopTo1; j++)
                        {
                            if (selectedRng.Cells[i, j].value is not null)
                            {
                                ValueFlag = "NotEmpty";
                            }
                        }
                    }

                    // loop through each cells of a column of the selection and check if the column is empty or not.
                    // if all the cells of the column of that seleted range are blank then the "flag" variable remains Empty. If any of the cell of that column is non-empty then "flag" will be "NotEmpty"
                    // if flag is "Empty" then the respective cells of that column of the selection will be deleted.
                    // Note that any cells of the same column that is outside the selection will be deleted even if it is empty
                    // after checking a column, the "flag" variable resets to "Empty"
                    for (int i = selectedRng.Columns.Count; i >= 1; i -= 1)
                    {
                        flag = "Empty";
                        for (int j = selectedRng.Rows.Count; j >= 1; j -= 1)
                        {
                            if (selectedRng.Cells[j, i].value is not null)
                            {

                                flag = "NotEmpty";

                            }

                        }


                        if (flag == "Empty")
                        {

                            worksheet.get_Range(selectedRng.Cells[1, i], selectedRng.Cells[selectedRng.Rows.Count, i]).Delete(XlDeleteShiftDirection.xlShiftToLeft);
                            blankColCount += 1;

                        }

                    }

                    // if no blank columns are found in a sheet then go to the "nextloop" section and skip the lines in between
                    if (blankColCount == 0)
                    {
                        goto break1;
                    }

                    // valueFlag is "Empty" means the entire selection is blank
                    // so the msgbox will be shown  and then exit sub                
                    if (ValueFlag == "Empty")
                    {
                        Interaction.MsgBox(blankColCount + " Column(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO");
                        return;
                    }

                    selectedRng.Cells[1, 1].select();
break1:
                    ;

                    // displays a msgbox that shows how many columns are deleted
                    Interaction.MsgBox(blankColCount + " Column(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO");
                }




                // user selected multiple disjoint ranges
                else
                {
                    // an array named "arrRng" is used to separately store all  the addresses of the selection 
                    string[] arrRng = Strings.Split(selectedRng.get_Address(), ",");

                    // loop through each address from the selection and check if any range is a single cell or not. If so, then the following warning will pop up.
                    // Then exit sub.
                    for (int i = 0, loopTo2 = Information.UBound(arrRng); i <= loopTo2; i++)
                    {
                        selectedRng = worksheet.get_Range(arrRng[i]);
                        if (selectedRng.Rows.Count == 1 & selectedRng.Columns.Count == 1)
                        {
                            Interaction.MsgBox("This Add-in doesn't work for single cell. Please select a Range and try again!", MsgBoxStyle.Exclamation, "Warning");
                            return;
                        }
                    }


                    // loops through the entire selection and if the entire selection is empty or not
                    // if the entire selection is blank then valueFlag will become "NotEmpty"  
                    for (int i = 0, loopTo3 = Information.UBound(arrRng); i <= loopTo3; i++)
                    {
                        for (int j = 1, loopTo4 = selectedRng.Rows.Count; j <= loopTo4; j++)
                        {
                            for (int k = 1, loopTo5 = selectedRng.Columns.Count; k <= loopTo5; k++)
                            {
                                if (selectedRng.Cells[j, k].value is not null)
                                {
                                    ValueFlag = "NotEmpty";
                                }
                            }
                        }
                    }


                    // loop through each range of the selection and remove blank columns
                    for (int i = 0, loopTo6 = Information.UBound(arrRng); i <= loopTo6; i++)
                    {

                        selectedRng = worksheet.get_Range(arrRng[i]);

                        // loop through each cells of a column of the selection and check if the column is empty or not.
                        // if all the cells of the column of that seleted range are blank then the "flag" variable remains "Empty". If any of the cell of that column is non-empty then "flag" will be "NotEmpty"
                        // if flag is "Empty" then the respective cells of that column of the selection will be deleted.
                        // Note that any cells of the same column that is outside the selection will be deleted even if it is empty
                        // after checking a column, the "flag" variable resets to "Empty"
                        for (int k = selectedRng.Columns.Count; k >= 1; k -= 1)
                        {
                            flag = "Empty";

                            for (int j = selectedRng.Rows.Count; j >= 1; j -= 1)
                            {

                                if (selectedRng.Cells[j, k].value is not null)
                                {

                                    flag = "NotEmpty";

                                }

                            }

                            if (flag == "Empty")
                            {

                                worksheet.get_Range(selectedRng.Cells[1, k], selectedRng.Cells[selectedRng.Rows.Count, k]).Delete(XlDeleteShiftDirection.xlShiftToLeft);
                                blankColCount += 1;

                            }

                        }

                    }

                    // if no blank columns are found in a sheet then go to the "nextloop" section and skip the lines in between
                    if (blankColCount == 0)
                    {
                        goto break2;
                    }

                    // valueFlag is "Empty" means the entire selection is blank
                    // so the msgbox will be shown  and then exit sub                
                    if (ValueFlag == "Empty")
                    {
                        Interaction.MsgBox(blankColCount + " Column(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO");
                        return;
                    }

                    selectedRng.Cells[1, 1].select();

break2:
                    ;

                    // displays a msgbox that shows how many columns are deleted
                    Interaction.MsgBox(blankColCount + " Column(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO");

                }
            }


            catch (Exception ex)
            {

            }


        }

        /// <summary>
    /// removes all blank columns from Active Sheet
    /// </summary>
        private void Button34_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {

                Excel.Application excelApp;
                Excel.Workbook workbook;
                Excel.Worksheet worksheet;
                Range selectedRng;
                int blankColCount = 0;
                string flag = "Empty";

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                selectedRng = (Range)excelApp.Selection;

                // use the UsedRange method to find the address of the range used in the active sheet
                // use split function to get 2nd portion of the range which is the last cell of the used range
                // Use this addrees to to find row and column number of last cell
                string[] lastCell;
                int lastColNum;
                int lastRowNum;

                lastCell = worksheet.UsedRange.get_Address().Split(':');
                lastColNum = worksheet.get_Range(lastCell[1]).Column;
                lastRowNum = worksheet.get_Range(lastCell[1]).Row;


                // loop through each cells of a column of the active sheet and check if the column is empty or not.
                // if all the cells of the column of that seleted range are blank then the "flag" variable remains "Empty". If any of the cell of that column is non-empty then "flag" will be "NotEmpty"
                // if flag is "Empty" then the respective column of the active sheet will be deleted.
                // after checking a column, the "flag" variable resets to "Empty"
                for (int i = lastColNum; i >= 1; i -= 1)
                {
                    flag = "Empty";
                    for (int j = lastRowNum; j >= 1; j -= 1)
                    {
                        if (worksheet.Cells[j, i].value is not null)
                        {

                            flag = "NotEmpty";

                        }

                    }

                    if (flag == "Empty")
                    {

                        worksheet.Cells[1, i].entirecolumn.delete();

                        blankColCount += 1;

                    }
                }

                // if no blank columns are found in a sheet then go to the "nextloop" section and skip the lines in between
                if (blankColCount == 0)
                {
                    goto break;
                }

                worksheet.Cells[1, 1].select();

break:
                ;

                // displays a msgbox that shows how many columns are deleted
                Interaction.MsgBox(blankColCount + " Column(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO");
            }

            catch (Exception ex)
            {

            }


        }


        /// <summary>
    /// removes blank columns from the selected worksheets
    /// </summary>

        private void Button35_Click(object sender, RibbonControlEventArgs e)
        {

            try
            {


                Excel.Application excelApp;
                Excel.Workbook workbook;
                Excel.Worksheet worksheet;
                string confirmationMsg = "";
                int blankColCount;
                int i = 0;
                string flag = "NotEmpty";


                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;

                // takes the sheet names of the selected worksheets
                var selectedSheets = excelApp.ActiveWindow.SelectedSheets;
                string sheetName = "";

                // loops through each selected worksheet and concatenate all the sheet names togehter in the "sheetName" variable
                // then Right function removes the leading comma (,) from the "sheetName" variable
                foreach (Excel.Worksheet sheet in selectedSheets)
                    sheetName = sheetName + "," + sheet.Name;
                sheetName = Strings.Right(sheetName, Strings.Len(sheetName) - 1);

                // new array (arrSheetName) stores all the selected sheet names separately
                string[] arrSheetName = Strings.Split(sheetName, ",");


                // loops through each selected sheet name from the "arrSheetName" array
                // "worksheet" variable takes the sheets name from the array and makes it active worksheet
                // each time a new sheet is taken from the slected sheets, "blankColCount" resets to 0
                var loopTo = Information.UBound(arrSheetName);
                for (i = 0; i <= loopTo; i++)
                {
                    blankColCount = 0;
                    worksheet = (Excel.Worksheet)workbook.Sheets[arrSheetName[i]];
                    worksheet.Activate();


                    // use the UsedRange method to find the address of the range used in the active sheet
                    // use split function to get 2nd portion of the range which is the last cell of the used range
                    // Use this addrees to to find row and column number of last cell
                    string[] lastCell;
                    int lastRowNum;
                    int lastColNum;

                    lastCell = worksheet.UsedRange.get_Address().Split(':');
                    lastRowNum = worksheet.get_Range(lastCell[1]).Row;
                    lastColNum = worksheet.get_Range(lastCell[1]).Column;

                    // loop through each cells of a column of the active sheet and check if the column is empty or not.
                    // if all the cells of the column of that seleted range are blank then the "flag" variable remains "Empty". If any of the cell of that column is non-empty then "flag" will be "NotEmpty"
                    // if flag is "Empty" then the respective column of the active sheet will be deleted.
                    // after checking a column, the "flag" variable resets to "Empty"
                    for (int j = lastColNum; j >= 1; j -= 1)
                    {
                        flag = "Empty";
                        for (int k = lastRowNum; k >= 1; k -= 1)
                        {
                            if (worksheet.Cells[k, j].value is not null)
                            {

                                flag = "NotEmpty";

                            }

                        }

                        if (flag == "Empty")
                        {

                            worksheet.Cells[1, j].entirecolumn.delete();

                            blankColCount += 1;

                        }
                    }

                    // if no blank columns are found in a sheet then go to the "nextloop" section and skip the lines in between
                    if (blankColCount == 0)
                    {
                        goto nextloop;
                    }


nextloop:
                    ;

                    // stores information about how many columns deleted from which sheet
                    confirmationMsg = confirmationMsg + blankColCount + " Column(s) are deleted from " + arrSheetName[i] + Microsoft.VisualBasic.Constants.vbCrLf;

                }

                // finally this msgBox is shown
                Interaction.MsgBox(confirmationMsg, MsgBoxStyle.Information, "SOFTEKO");
            }

            catch (Exception ex)
            {

            }



        }


        /// <summary>
    /// removes blank columns from all worksheets from the active workbook
    /// </summary>

        private void Button36_Click(object sender, RibbonControlEventArgs e)
        {

            try
            {

                Excel.Application excelApp;
                Excel.Workbook workbook;
                Excel.Worksheet worksheet;
                string confirmationMsg = "";
                int blankColCount;
                int i = 0;
                string flag = "Empty";


                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;

                // takes the sheet names of all worksheets of the workbook
                var selectedSheets = excelApp.Sheets;
                string sheetName = "";

                // loops through each selected worksheet and concatenate all the sheet names togehter in the "sheetName" variable
                // then Right function removes the leading comma (,) from the "sheetName" variable
                foreach (Excel.Worksheet sheet in selectedSheets)
                    sheetName = sheetName + "," + sheet.Name;
                sheetName = Strings.Right(sheetName, Strings.Len(sheetName) - 1);

                // new array (arrSheetName) stores all the sheet names separately
                string[] arrSheetName = Strings.Split(sheetName, ",");


                // loops through each sheet name from the "arrSheetName" array
                // "worksheet" variable takes the sheet names from the array and makes it active worksheet
                // each time a new sheet is taken by "worksheet" variable, "blankColList" and "blankColCount" resets to 0
                var loopTo = Information.UBound(arrSheetName);
                for (i = 0; i <= loopTo; i++)
                {
                    blankColCount = 0;
                    worksheet = (Excel.Worksheet)workbook.Sheets[arrSheetName[i]];
                    worksheet.Activate();


                    // use the UsedRange method to find the address of the range used in the active sheet
                    // use split function to get 2nd portion of the range which is the last cell of the used range
                    // Use this addrees to to find column number of last cell
                    string[] lastCell;
                    int lastRowNum;
                    int lastColNum;

                    lastCell = worksheet.UsedRange.get_Address().Split(':');
                    lastRowNum = worksheet.get_Range(lastCell[1]).Row;
                    lastColNum = worksheet.get_Range(lastCell[1]).Column;

                    // loop through each cells of a column of the active sheet and check if the column is empty or not.
                    // if all the cells of the column of that seleted range are blank then the "flag" variable remains "Empty". If any of the cell of that column is non-empty then "flag" will be "NotEmpty"
                    // if flag is "Empty" then the respective column of the active sheet will be deleted.
                    // after checking a column, the "flag" variable resets to "Empty"
                    for (int j = lastColNum; j >= 1; j -= 1)
                    {
                        flag = "Empty";
                        for (int k = lastRowNum; k >= 1; k -= 1)
                        {
                            if (worksheet.Cells[k, j].value is not null)
                            {

                                flag = "NotEmpty";

                            }

                        }

                        if (flag == "Empty")
                        {

                            worksheet.Cells[1, j].entirecolumn.delete();

                            blankColCount += 1;

                        }
                    }

                    // if no blank columns are found in a sheet then go to the "nextloop" section and skip the lines in between
                    if (blankColCount == 0)
                    {
                        goto nextloop;
                    }

nextloop:
                    ;

                    // stores information about how many columns deleted from which sheet
                    confirmationMsg = confirmationMsg + blankColCount + " Column(s) are deleted from " + arrSheetName[i] + Microsoft.VisualBasic.Constants.vbCrLf;

                }

                // finally this msgBox is shown
                Interaction.MsgBox(confirmationMsg, MsgBoxStyle.Information, "SOFTEKO");
            }

            catch (Exception ex)
            {

            }


        }


        /// <summary>
    /// removes blank rows from selected range of the active worksheet
    /// </summary>
        public static List<string> blank_Cell_Address_List = new List<string>();
        public static int blankRowCount;
        public static string captiontxt;

        private void Button37_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp;
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;

            excelApp = Globals.ThisAddIn.Application;
            workbook = excelApp.ActiveWorkbook;
            worksheet = (Excel.Worksheet)excelApp.ActiveSheet;

            emptyRows_selectedRng();








            var startTime = DateTime.Now;
            int totalItems = blank_Cell_Address_List.Count; // Total number of items to process
            int itemsProcessed = 0;
            TimeSpan remainingTime;

            var form = new FormProgressBar();
            form.Show();
            form.Activate();

            var Label = new System.Windows.Forms.Label();
            Label.Text = "Estimated remaining time: ";

            Label.Location = new System.Drawing.Point(12, 20);
            Label.Height = 15;
            Label.Width = 450;
            Label.TextAlign = ContentAlignment.TopLeft;
            form.Controls.Add(Label);


            for (int k = 0, loopTo = blank_Cell_Address_List.Count - 1; k <= loopTo; k++)
            {

                // Process each item
                worksheet.get_Range(blank_Cell_Address_List[k]).Delete(XlDeleteShiftDirection.xlShiftUp);

                itemsProcessed += 1;



                // For i = 1 To itemsProcessed
                // FormProgressBar.proBar.Value = (i / itemsProcessed) * 100
                // Label.Text = "Estimated remaining time: " + remainingTime.ToString()
                // Next

                // Calculate elapsed time


                // Estimate remaining time
                if (itemsProcessed > 0)
                {
                    var elapsedTime = DateTime.Now - startTime;
                    var totalEstimatedTime = TimeSpan.FromTicks((long)Math.Round(elapsedTime.Ticks * totalItems / (double)itemsProcessed));
                    // Dim elapsedTime As TimeSpan = DateTime.Now - startTime
                    remainingTime = totalEstimatedTime - elapsedTime;
                    string formattedTime = string.Format("{0:00}:{1:00}", (int)Math.Round(Math.Truncate(remainingTime.TotalMinutes)), remainingTime.Seconds);


                    Label.Text = "Estimated remaining time: " + formattedTime;
                    FormProgressBar.proBar.Value = (int)Math.Round(itemsProcessed / (double)totalItems * 100d);

                    // Display or use the remaining time
                    // Console.WriteLine("Estimated remaining time: " + remainingTime.ToString())

                    // captiontxt = "Estimated remaining time: " + remainingTime.ToString()
                }
            }

            // Dim form As New FormProgressBar
            // form.Show()
            // form.Activate()

            // Dim Label As New System.Windows.Forms.Label

            // Label.Text = "Estimated remaining time: " + remainingTime.ToString()
            // Label.Location = New System.Drawing.Point(12, 20)
            // Label.Height = 15
            // Label.Width = 450
            // Label.TextAlign = ContentAlignment.TopLeft
            // form.Controls.Add(Label)

            // For i = 1 To itemsProcessed
            // FormProgressBar.proBar.Value = (i / itemsProcessed) * 100
            // Label.Text = "Estimated remaining time: " + remainingTime.ToString()
            // Next






            // Try
            // Dim excelApp As Excel.Application
            // Dim workbook As Excel.Workbook
            // Dim worksheet As Excel.Worksheet
            // Dim selectedRng As Excel.Range
            // Dim blankRowCount As Integer = 0
            // Dim flag As String = "Empty"
            // Dim ValueFlag As String = "Empty"
            // Dim answer As MsgBoxResult
            // Dim blankCells As Range
            // Dim columnPattern As String = "^(\$?[A-Z]+)(:)(\$?[A-Z]+)$"
            // Dim rowPattern As String = "^(\$?[0-9]+)(:)(\$?[0-9]+)$"
            // Dim isMatchRow, isMatchColumn As Boolean
            // Dim blankCellRowNumber As New List(Of Integer)
            // Dim startTime, endTime As DateTime
            // Dim runTime As TimeSpan

            // startTime = Now
            // excelApp = Globals.ThisAddIn.Application
            // workbook = excelApp.ActiveWorkbook
            // worksheet = workbook.ActiveSheet
            // selectedRng = excelApp.Selection


            // '"rngCount" variable indicates the number of ranges in the users' selection.
            // '0 means a single continuous selection
            // '> 0 means multiple disjoint range selection
            // Dim rngCount As Integer
            // rngCount = 0

            // For Each c As Char In selectedRng.Address

            // If c = "," Then
            // rngCount = rngCount + 1
            // End If

            // Next



            // 'if user select a single continuous range "rngCount" will be 0
            // If rngCount = 0 Then

            // isMatchRow = Regex.IsMatch(UCase(selectedRng.Address), rowPattern)


            // 'columnPattern = "^(\$?[A-Z]+)(:)(\$?[A-Z]+)$"
            // isMatchColumn = Regex.IsMatch(UCase(selectedRng.Address), columnPattern)

            // 'checks if the user selected an entire column or not, if yes then following block will excecute
            // If isMatchColumn = True And isMatchRow = False Then

            // 'if there is no blank cell in the entire column then an exception will be thrown and 0 cells will be deleted
            // Try
            // blankCells = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks)
            // blankRowCount = blankCells.Cells.Count

            // answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
            // If answer = MsgBoxResult.No Then
            // Exit Sub
            // Else
            // blankCells.Delete(Excel.XlDeleteShiftDirection.xlShiftUp)

            // 'displays a msgbox that shows how many rows are deleted
            // MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

            // End If

            // Catch ex As Exception

            // MsgBox(0 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
            // Exit Sub

            // End Try

            // 'if user selects an entire row as selection then following block will execute
            // ElseIf isMatchColumn = False And isMatchRow = True Then

            // blankRowCount = 0

            // For i = selectedRng.Rows.Count To 1 Step -1
            // If excelApp.WorksheetFunction.CountA(selectedRng.Rows(i)) = 0 Then
            // blankRowCount += 1
            // End If
            // Next
            // If blankRowCount = 0 Then
            // MsgBox(0 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
            // Exit Sub
            // End If

            // answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
            // If answer = MsgBoxResult.No Then
            // Exit Sub
            // Else

            // blankRowCount = 0

            // For i = selectedRng.Rows.Count To 1 Step -1

            // 'checks if the entire row is empty or not
            // If excelApp.WorksheetFunction.CountA(selectedRng.Rows(i)) = 0 Then
            // blankRowCount += 1
            // selectedRng.Rows(i).delete
            // End If
            // Next

            // selectedRng.Cells(1, 1).end(XlDirection.xlToRight).select
            // 'displays a msgbox that shows how many rows are deleted
            // MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

            // End If

            // 'if user didnt select an entire column or an entire row then following block will excecute
            // Else

            // 'if user select only a single cell the following warning will pop up and exit from the code 
            // If selectedRng.Rows.Count = 1 And selectedRng.Columns.Count = 1 Then
            // If selectedRng.Cells(1, 1).value Is Nothing Then
            // answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
            // If answer = MsgBoxResult.No Then
            // Exit Sub
            // Else
            // selectedRng.Cells(1, 1).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            // MsgBox(1 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
            // Exit Sub
            // End If
            // Else
            // MsgBox(0 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
            // Exit Sub
            // End If
            // End If



            // 'if there is no blank cell in the selection then an exception will be thrown and 0 cells will be deleted
            // Try
            // blankCells = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks)

            // Dim arr_Blank_Cell_Address() As String = Split(blankCells.Address, ",")

            // answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
            // If answer = MsgBoxResult.No Then
            // Exit Sub
            // Else

            // For i = 0 To UBound(arr_Blank_Cell_Address)

            // If worksheet.Range(arr_Blank_Cell_Address(i)).Columns.Count = selectedRng.Columns.Count Then
            // worksheet.Range(arr_Blank_Cell_Address(i)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            // blankRowCount += worksheet.Range(arr_Blank_Cell_Address(i)).Rows.Count
            // Else
            // Continue For
            // End If
            // Next

            // selectedRng.Cells(1, 1).select

            // 'displays a msgbox that shows how many rows are deleted
            // MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

            // End If

            // Catch ex As Exception

            // MsgBox(0 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
            // Exit Sub

            // End Try


            // End If




            // '                Dim pattern As String = "^(\$?[A-Z]+)(:)(\$?[A-Z]+)$"
            // '                Dim isMatch As Boolean = Regex.IsMatch(UCase(selectedRng.Address), pattern)
            // '                Dim outStr As String = ""
            // '                If isMatch = True Then

            // '                    answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
            // '                    If answer = MsgBoxResult.No Then
            // '                        Exit Sub
            // '                    Else

            // '                        Dim blankCells As Range = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks)
            // '                        blankRowCount = blankCells.Cells.Count
            // '                        blankCells.Delete(Excel.XlDeleteShiftDirection.xlShiftUp)

            // '                        'displays a msgbox that shows how many rows are deleted
            // '                        MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
            // '                    End If


            // '                Else

            // '                    'store the first row number and column number of the selected range in 2 variables
            // '                    Dim firstRowNum = selectedRng.Row
            // '                    Dim firstColNum = selectedRng.Column

            // '                    'if user select only a single cell the following warning will pop up and exit from the code 
            // '                    If selectedRng.Rows.Count = 1 And selectedRng.Columns.Count = 1 Then
            // '                        If selectedRng.Cells(1, 1).value Is Nothing Then
            // '                            answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
            // '                            If answer = MsgBoxResult.No Then
            // '                                Exit Sub
            // '                            Else
            // '                                selectedRng.Cells(1, 1).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            // '                                MsgBox(1 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
            // '                                Exit Sub
            // '                            End If
            // '                            'MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.Exclamation, "Warning")
            // '                        End If
            // '                    End If


            // '                    'loops through the entire selection and if the entire selection is empty or not
            // '                    'if the entire selection is blank then valueFlag will become "Empty"
            // '                    For i = 1 To selectedRng.Rows.Count
            // '                        For j = 1 To selectedRng.Columns.Count
            // '                            If Not selectedRng.Cells(i, j).value Is Nothing Then
            // '                                ValueFlag = "NotEmpty"
            // '                            End If
            // '                        Next
            // '                    Next


            // '                    'check if the selection has any deletable row
            // '                    'if there is any then a msgbox will pop up to warn the user that the data will move up after deleting row
            // '                    'if user selects yes then the rows will be deleted as usual
            // '                    'if user selects no then it will exit the sub and nothing will happen
            // '                    For i = selectedRng.Rows.Count To 1 Step -1
            // '                        flag = "Empty"
            // '                        For j = selectedRng.Columns.Count To 1 Step -1
            // '                            If Not selectedRng.Cells(i, j).value Is Nothing Then

            // '                                flag = "NotEmpty"

            // '                            End If

            // '                        Next


            // '                        If flag = "Empty" Then

            // '                            answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
            // '                            If answer = MsgBoxResult.No Then
            // '                                Exit Sub
            // '                            End If
            // '                            Exit For

            // '                        End If

            // '                    Next



            // '                    'loop through each cells of a row of the selection and check if the row is empty or not.
            // '                    'if all the cells of the row of that seleted range are blank then the "flag" variable remains Empty. If any of the cell of that row is non-empty then "flag" will be "NotEmpty"
            // '                    'if flag is "Empty" the blank row is deleted and increase the value of "blankRowCount" by 1
            // '                    'after checking a row, the "flag" variable resets to "Empty"
            // '                    For i = selectedRng.Rows.Count To 1 Step -1
            // '                        flag = "Empty"
            // '                        For j = selectedRng.Columns.Count To 1 Step -1
            // '                            If Not selectedRng.Cells(i, j).value Is Nothing Then

            // '                                flag = "NotEmpty"

            // '                            End If

            // '                        Next


            // '                        If flag = "Empty" Then

            // '                            worksheet.Range(selectedRng.Cells(i, 1), selectedRng.Cells(i, selectedRng.Columns.Count)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            // '                            blankRowCount = blankRowCount + 1

            // '                        End If

            // '                    Next



            // '                    'if no blank rows are found in a sheet then go to the "break1" section and skip the lines in between
            // '                    If blankRowCount = 0 Then
            // '                        GoTo break1
            // '                    End If


            // '                    'valueFlag is "Empty" means the entire selection is blank
            // '                    'so the msgbox will be shown  and then exit sub   
            // '                    If ValueFlag = "Empty" Then
            // '                        worksheet.Cells(firstRowNum, firstColNum).select
            // '                        MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
            // '                        Exit Sub
            // '                    End If

            // '                    selectedRng.Cells(1, 1).select

            // 'break1:
            // '                    'displays a msgbox that shows how many rows are deleted
            // '                    MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

            // '                End If


            // 'user selected multiple disjoint ranges

            // Else

            // Dim arrRng As String() = Split(selectedRng.Address, ",")

            // Dim totalBlankRowCount As Integer = 0

            // answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
            // If answer = MsgBoxResult.No Then
            // Exit Sub
            // Else

            // For i = 0 To UBound(arrRng)
            // selectedRng = worksheet.Range(arrRng(i))


            // isMatchRow = Regex.IsMatch(UCase(selectedRng.Address), rowPattern)


            // 'columnPattern = "^(\$?[A-Z]+)(:)(\$?[A-Z]+)$"
            // isMatchColumn = Regex.IsMatch(UCase(selectedRng.Address), columnPattern)

            // 'checks if the user selected an entire column or not, if yes then following block will excecute
            // If isMatchColumn = True And isMatchRow = False Then

            // 'if there is no blank cell in the entire column then an exception will be thrown 
            // Try
            // blankCells = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks)
            // blankRowCount = blankCells.Cells.Count

            // blankCells.Delete(Excel.XlDeleteShiftDirection.xlShiftUp)

            // Catch ex As Exception

            // GoTo nextRange

            // End Try


            // 'if user selects an entire row as selection then following block will execute
            // ElseIf isMatchColumn = False And isMatchRow = True Then


            // blankRowCount = 0

            // For k = selectedRng.Rows.Count To 1 Step -1

            // 'checks if the entire row is empty or not
            // If excelApp.WorksheetFunction.CountA(selectedRng.Rows(k)) = 0 Then
            // blankRowCount += 1
            // selectedRng.Rows(k).delete
            // End If
            // Next



            // 'if user didnt select an entire column or an entire row then following block will excecute
            // Else
            // blankRowCount = 0
            // 'if user select only a single cell the following warning will pop up and exit from the code 
            // If selectedRng.Rows.Count = 1 And selectedRng.Columns.Count = 1 Then
            // If selectedRng.Cells(1, 1).value Is Nothing Then
            // selectedRng.Cells(1, 1).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            // blankRowCount += 1
            // GoTo nextRange
            // Else
            // GoTo nextRange
            // End If
            // End If



            // 'if there is no blank cell in the selection then an exception will be thrown and 0 cells will be deleted
            // Try
            // blankCells = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks)

            // Dim arr_Blank_Cell_Address() As String = Split(blankCells.Address, ",")

            // For k = 0 To UBound(arr_Blank_Cell_Address)

            // If worksheet.Range(arr_Blank_Cell_Address(k)).Columns.Count = selectedRng.Columns.Count Then
            // worksheet.Range(arr_Blank_Cell_Address(k)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            // blankRowCount += worksheet.Range(arr_Blank_Cell_Address(k)).Rows.Count
            // Else
            // GoTo nextRange
            // End If
            // Next

            // Catch ex As Exception

            // End Try

            // End If
            // nextRange:
            // totalBlankRowCount += blankRowCount
            // Next

            // End If

            // MsgBox(totalBlankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

            // End If


            // endTime = Now

            // runTime = endTime.Subtract(startTime)

            // If runTime.TotalSeconds > 5 Then

            // Dim form As New FormProgressBar
            // form.Show()
            // 'Dim startTime As DateTime = DateTime.Now
            // Dim totalItems As Integer = blankRowCount ' Total number of items to process
            // Dim itemsProcessed As Integer = 0

            // For k As Integer = 1 To totalItems
            // ' Process each item
            // Application.DoEvents()

            // itemsProcessed += 1

            // ' Calculate elapsed time
            // Dim elapsedTime As TimeSpan = DateTime.Now - startTime

            // ' Estimate remaining time
            // If itemsProcessed > 0 Then
            // Dim totalEstimatedTime As TimeSpan = TimeSpan.FromTicks(elapsedTime.Ticks * totalItems / itemsProcessed)
            // Dim remainingTime As TimeSpan = totalEstimatedTime - elapsedTime

            // ' Display or use the remaining time
            // Console.WriteLine("Estimated remaining time: " + remainingTime.ToString())



            // End If
            // Next


            // End If

            // Catch ex As Exception

            // End Try



            // Else
            // 'an array named "arrRng" is used to separately store all  the addresses of the selection 
            // 'Dim arrRng As String() = Split(selectedRng.Address, ",")

            // 'Dim totalBlankRowCount As Integer = 0
            // Try
            // For i = 0 To UBound(arrRng)
            // blankCells = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks)
            // If worksheet.Range(arrRng(i)).Columns.Count = selectedRng.Columns.Count Then
            // blankRowCount = worksheet.Range(arrRng(i)).Rows.Count
            // totalBlankRowCount += blankRowCount
            // End If

            // Next
            // 'MsgBox(totalBlankRowCount)


            // Catch ex As Exception
            // MsgBox(0 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
            // Exit Sub
            // End Try

            // answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
            // If answer = MsgBoxResult.No Then

            // Exit Sub

            // Else

            // totalBlankRowCount = 0

            // 'loop through each address from the selection and check if any range is a single cell or not. If so, then the following warning will pop up.
            // 'Then exit sub.
            // For i = 0 To UBound(arrRng)

            // selectedRng = worksheet.Range(arrRng(i))

            // columnPattern = "^(\$?[A-Z]+)(:)(\$?[A-Z]+)$"
            // isMatchColumn = Regex.IsMatch(UCase(selectedRng.Address), columnPattern)

            // 'checks if the user selected an entire column or not, if yes then following block will excecute
            // If isMatchColumn = True Then


            // blankCells = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks)
            // blankRowCount = blankCells.Cells.Count

            // blankCells.Delete(Excel.XlDeleteShiftDirection.xlShiftUp)

            // 'if user didnt select an entire column then following block will excecute
            // Else

            // 'if user select only a single cell the following warning will pop up and exit from the code 
            // If selectedRng.Rows.Count = 1 And selectedRng.Columns.Count = 1 Then
            // If selectedRng.Cells(1, 1).value Is Nothing Then
            // selectedRng.Cells(1, 1).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            // blankRowCount = 1
            // End If
            // Else



            // 'if there is no blank cell in the selection then an exception will be thrown and 0 cells will be deleted
            // Try
            // blankCells = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks)

            // Dim arr_Blank_Cell_Address() As String = Split(blankCells.Address, ",")
            // For j = 0 To UBound(arr_Blank_Cell_Address)
            // If worksheet.Range(arr_Blank_Cell_Address(j)).Columns.Count = selectedRng.Columns.Count Then
            // worksheet.Range(arr_Blank_Cell_Address(j)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            // blankRowCount += worksheet.Range(arr_Blank_Cell_Address(j)).Rows.Count
            // Else
            // Continue For
            // End If
            // Next

            // Catch ex As Exception
            // blankRowCount = 0
            // End Try
            // End If


            // End If

            // totalBlankRowCount += blankRowCount


            // Next
            // MsgBox(totalBlankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

            // End If


            // Exit Sub





            // If selectedRng.Rows.Count = 1 And selectedRng.Columns.Count = 1 Then
            // MsgBox("This Add-in doesn't work for single cell. Please select a Range and try again!", MsgBoxStyle.Exclamation, "Warning")
            // Exit Sub
            // End If

            // loops through the entire selection and if the entire selection is empty or not
            // if the entire selection is blank then valueFlag will become "NotEmpty"
            // For i = 0 To UBound(arrRng)
            // selectedRng = worksheet.Range(arrRng(i))
            // For j = 1 To selectedRng.Rows.Count
            // For k = 1 To selectedRng.Columns.Count
            // If Not selectedRng.Cells(j, k).value Is Nothing Then
            // ValueFlag = "NotEmpty"
            // End If
            // Next
            // Next
            // Next

            // 'loop through each range of the selection and remove blank rows
            // For i = 0 To UBound(arrRng)

            // selectedRng = worksheet.Range(arrRng(i))

            // 'loop through each cells of a row of the selection and check if the row is empty or not.
            // 'if all the cells of the row of that seleted range are blank then the "flag" variable remains Empty. If any of the cell of that row is non-empty then "flag" will be "NotEmpty"
            // 'if flag is "Empty" the blank row is deleted and increase the value of "blankRowCount" by 1
            // 'after checking a row, the "flag" variable resets to "Empty"
            // For j = selectedRng.Rows.Count To 1 Step -1
            // flag = "Empty"
            // For k = selectedRng.Columns.Count To 1 Step -1
            // If Not selectedRng.Cells(j, k).value Is Nothing Then

            // flag = "NotEmpty"

            // End If

            // Next

            // If flag = "Empty" Then

            // worksheet.Range(selectedRng.Cells(j, 1), selectedRng.Cells(j, selectedRng.Columns.Count)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            // blankRowCount = blankRowCount + 1

            // End If

            // Next

            // Next

            // 'if no blank rows are found in a sheet then go to the "break2" section and skip the lines in between
            // If blankRowCount = 0 Then
            // GoTo break2
            // End If

            // 'valueFlag is "Empty" means the entire selection is blank
            // 'so the msgbox will be shown  and then exit sub  
            // If ValueFlag = "Empty" Then
            // MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

            // Exit Sub
            // End If

            // selectedRng.Cells(1, 1).select

            // break2:
            // 'displays a msgbox that shows how many rows are deleted
            // MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")



        }

        private void emptyRows_selectedRng()
        {

            try
            {
                Excel.Application excelApp;
                Excel.Workbook workbook;
                Excel.Worksheet worksheet;
                Range selectedRng;
                // Dim blankRowCount As Integer = 0
                string flag = "Empty";
                string ValueFlag = "Empty";
                MsgBoxResult answer;
                Range blankCells;
                string columnPattern = @"^(\$?[A-Z]+)(:)(\$?[A-Z]+)$";
                string rowPattern = @"^(\$?[0-9]+)(:)(\$?[0-9]+)$";
                bool isMatchRow, isMatchColumn;
                var blankCellRowNumber = new List<int>();
                // Dim startTime, endTime As DateTime
                // Dim runTime As TimeSpan

                // startTime = Now
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                selectedRng = (Range)excelApp.Selection;

                blank_Cell_Address_List.Clear();

                string[] arrRng = Strings.Split(selectedRng.get_Address(), ",");
                string[] arrBlankCellAddress;
                int totalBlankRowCount = 0;

                for (int i = 0, loopTo = Information.UBound(arrRng); i <= loopTo; i++)
                {
                    selectedRng = worksheet.get_Range(arrRng[i]);
                    blankCells = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks);
                    arrBlankCellAddress = Strings.Split(blankCells.get_Address(), ",");

                    for (int j = 0, loopTo1 = Information.UBound(arrBlankCellAddress); j <= loopTo1; j++)


                        blank_Cell_Address_List.Add(arrBlankCellAddress[j]);

                }
            }

            // For i = 0 To blank_Cell_Address_List.Count
            // MsgBox(blank_Cell_Address_List(i))
            // Next















            // '"rngCount" variable indicates the number of ranges in the users' selection.
            // '0 means a single continuous selection
            // '> 0 means multiple disjoint range selection
            // Dim rngCount As Integer
            // rngCount = 0

            // For Each c As Char In selectedRng.Address

            // If c = "," Then
            // rngCount = rngCount + 1
            // End If

            // Next

            // blankRowCount = 0

            // 'if user select a single continuous range "rngCount" will be 0
            // If rngCount = 0 Then

            // isMatchRow = Regex.IsMatch(UCase(selectedRng.Address), rowPattern)


            // 'columnPattern = "^(\$?[A-Z]+)(:)(\$?[A-Z]+)$"
            // isMatchColumn = Regex.IsMatch(UCase(selectedRng.Address), columnPattern)

            // 'checks if the user selected an entire column or not, if yes then following block will excecute
            // If isMatchColumn = True And isMatchRow = False Then

            // 'if there is no blank cell in the entire column then an exception will be thrown and 0 cells will be deleted
            // Try
            // blankCells = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks)
            // blankRowCount = blankCells.Cells.Count

            // answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
            // If answer = MsgBoxResult.No Then
            // Exit Sub
            // Else
            // blankCells.Delete(Excel.XlDeleteShiftDirection.xlShiftUp)

            // 'displays a msgbox that shows how many rows are deleted
            // MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

            // End If

            // Catch ex As Exception

            // MsgBox(0 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
            // Exit Sub

            // End Try

            // 'if user selects an entire row as selection then following block will execute
            // ElseIf isMatchColumn = False And isMatchRow = True Then

            // blankRowCount = 0

            // For i = selectedRng.Rows.Count To 1 Step -1
            // If excelApp.WorksheetFunction.CountA(selectedRng.Rows(i)) = 0 Then
            // blankRowCount += 1
            // End If
            // Next
            // If blankRowCount = 0 Then
            // MsgBox(0 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
            // Exit Sub
            // End If

            // answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
            // If answer = MsgBoxResult.No Then
            // Exit Sub
            // Else

            // blankRowCount = 0

            // For i = selectedRng.Rows.Count To 1 Step -1

            // 'checks if the entire row is empty or not
            // If excelApp.WorksheetFunction.CountA(selectedRng.Rows(i)) = 0 Then
            // blankRowCount += 1
            // selectedRng.Rows(i).delete
            // End If
            // Next

            // selectedRng.Cells(1, 1).end(XlDirection.xlToRight).select
            // 'displays a msgbox that shows how many rows are deleted
            // MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

            // End If

            // 'if user didnt select an entire column or an entire row then following block will excecute
            // Else

            // 'if user select only a single cell the following warning will pop up and exit from the code 
            // If selectedRng.Rows.Count = 1 And selectedRng.Columns.Count = 1 Then
            // If selectedRng.Cells(1, 1).value Is Nothing Then
            // answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
            // If answer = MsgBoxResult.No Then
            // Exit Sub
            // Else
            // selectedRng.Cells(1, 1).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            // MsgBox(1 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
            // Exit Sub
            // End If
            // Else
            // MsgBox(0 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
            // Exit Sub
            // End If
            // End If



            // 'if there is no blank cell in the selection then an exception will be thrown and 0 cells will be deleted
            // Try
            // blankCells = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks)

            // Dim arr_Blank_Cell_Address() As String = Split(blankCells.Address, ",")

            // answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
            // If answer = MsgBoxResult.No Then
            // Exit Sub
            // Else

            // For i = 0 To UBound(arr_Blank_Cell_Address)

            // If worksheet.Range(arr_Blank_Cell_Address(i)).Columns.Count = selectedRng.Columns.Count Then
            // worksheet.Range(arr_Blank_Cell_Address(i)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            // blankRowCount += worksheet.Range(arr_Blank_Cell_Address(i)).Rows.Count
            // Else
            // Continue For
            // End If
            // Next

            // selectedRng.Cells(1, 1).select

            // 'displays a msgbox that shows how many rows are deleted
            // MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

            // End If

            // Catch ex As Exception

            // MsgBox(0 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
            // Exit Sub

            // End Try


            // End If


            // Else

            // Dim arrRng As String() = Split(selectedRng.Address, ",")

            // Dim totalBlankRowCount As Integer = 0

            // answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
            // If answer = MsgBoxResult.No Then
            // Exit Sub
            // Else

            // For i = 0 To UBound(arrRng)
            // selectedRng = worksheet.Range(arrRng(i))


            // isMatchRow = Regex.IsMatch(UCase(selectedRng.Address), rowPattern)


            // 'columnPattern = "^(\$?[A-Z]+)(:)(\$?[A-Z]+)$"
            // isMatchColumn = Regex.IsMatch(UCase(selectedRng.Address), columnPattern)

            // 'checks if the user selected an entire column or not, if yes then following block will excecute
            // If isMatchColumn = True And isMatchRow = False Then

            // 'if there is no blank cell in the entire column then an exception will be thrown 
            // Try
            // blankCells = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks)
            // blankRowCount = blankCells.Cells.Count

            // blankCells.Delete(Excel.XlDeleteShiftDirection.xlShiftUp)

            // Catch ex As Exception

            // GoTo nextRange

            // End Try


            // 'if user selects an entire row as selection then following block will execute
            // ElseIf isMatchColumn = False And isMatchRow = True Then


            // blankRowCount = 0

            // For k = selectedRng.Rows.Count To 1 Step -1

            // 'checks if the entire row is empty or not
            // If excelApp.WorksheetFunction.CountA(selectedRng.Rows(k)) = 0 Then
            // blankRowCount += 1
            // selectedRng.Rows(k).delete
            // End If
            // Next



            // 'if user didnt select an entire column or an entire row then following block will excecute
            // Else
            // blankRowCount = 0
            // 'if user select only a single cell the following warning will pop up and exit from the code 
            // If selectedRng.Rows.Count = 1 And selectedRng.Columns.Count = 1 Then
            // If selectedRng.Cells(1, 1).value Is Nothing Then
            // selectedRng.Cells(1, 1).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            // blankRowCount += 1
            // GoTo nextRange
            // Else
            // GoTo nextRange
            // End If
            // End If



            // 'if there is no blank cell in the selection then an exception will be thrown and 0 cells will be deleted
            // Try
            // blankCells = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks)

            // Dim arr_Blank_Cell_Address() As String = Split(blankCells.Address, ",")

            // For k = 0 To UBound(arr_Blank_Cell_Address)

            // If worksheet.Range(arr_Blank_Cell_Address(k)).Columns.Count = selectedRng.Columns.Count Then
            // worksheet.Range(arr_Blank_Cell_Address(k)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            // blankRowCount += worksheet.Range(arr_Blank_Cell_Address(k)).Rows.Count
            // Else
            // GoTo nextRange
            // End If
            // Next

            // Catch ex As Exception

            // End Try

            // End If
            // nextRange:
            // totalBlankRowCount += blankRowCount
            // Next

            // End If

            // MsgBox(totalBlankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

            // End If


            catch (Exception ex)
            {

            }

        }




        // Private Sub emptyRows_selectedRng()

        // Try
        // Dim excelApp As Excel.Application
        // Dim workbook As Excel.Workbook
        // Dim worksheet As Excel.Worksheet
        // Dim selectedRng As Excel.Range
        // 'Dim blankRowCount As Integer = 0
        // Dim flag As String = "Empty"
        // Dim ValueFlag As String = "Empty"
        // Dim answer As MsgBoxResult
        // Dim blankCells As Range
        // Dim columnPattern As String = "^(\$?[A-Z]+)(:)(\$?[A-Z]+)$"
        // Dim rowPattern As String = "^(\$?[0-9]+)(:)(\$?[0-9]+)$"
        // Dim isMatchRow, isMatchColumn As Boolean
        // Dim blankCellRowNumber As New List(Of Integer)
        // 'Dim startTime, endTime As DateTime
        // 'Dim runTime As TimeSpan

        // 'startTime = Now
        // excelApp = Globals.ThisAddIn.Application
        // workbook = excelApp.ActiveWorkbook
        // worksheet = workbook.ActiveSheet
        // selectedRng = excelApp.Selection


        // '"rngCount" variable indicates the number of ranges in the users' selection.
        // '0 means a single continuous selection
        // '> 0 means multiple disjoint range selection
        // Dim rngCount As Integer
        // rngCount = 0

        // For Each c As Char In selectedRng.Address

        // If c = "," Then
        // rngCount = rngCount + 1
        // End If

        // Next

        // blankRowCount = 0

        // 'if user select a single continuous range "rngCount" will be 0
        // If rngCount = 0 Then

        // isMatchRow = Regex.IsMatch(UCase(selectedRng.Address), rowPattern)


        // 'columnPattern = "^(\$?[A-Z]+)(:)(\$?[A-Z]+)$"
        // isMatchColumn = Regex.IsMatch(UCase(selectedRng.Address), columnPattern)

        // 'checks if the user selected an entire column or not, if yes then following block will excecute
        // If isMatchColumn = True And isMatchRow = False Then

        // 'if there is no blank cell in the entire column then an exception will be thrown and 0 cells will be deleted
        // Try
        // blankCells = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks)
        // blankRowCount = blankCells.Cells.Count

        // answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
        // If answer = MsgBoxResult.No Then
        // Exit Sub
        // Else
        // blankCells.Delete(Excel.XlDeleteShiftDirection.xlShiftUp)

        // 'displays a msgbox that shows how many rows are deleted
        // MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

        // End If

        // Catch ex As Exception

        // MsgBox(0 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
        // Exit Sub

        // End Try

        // 'if user selects an entire row as selection then following block will execute
        // ElseIf isMatchColumn = False And isMatchRow = True Then

        // blankRowCount = 0

        // For i = selectedRng.Rows.Count To 1 Step -1
        // If excelApp.WorksheetFunction.CountA(selectedRng.Rows(i)) = 0 Then
        // blankRowCount += 1
        // End If
        // Next
        // If blankRowCount = 0 Then
        // MsgBox(0 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
        // Exit Sub
        // End If

        // answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
        // If answer = MsgBoxResult.No Then
        // Exit Sub
        // Else

        // blankRowCount = 0

        // For i = selectedRng.Rows.Count To 1 Step -1

        // 'checks if the entire row is empty or not
        // If excelApp.WorksheetFunction.CountA(selectedRng.Rows(i)) = 0 Then
        // blankRowCount += 1
        // selectedRng.Rows(i).delete
        // End If
        // Next

        // selectedRng.Cells(1, 1).end(XlDirection.xlToRight).select
        // 'displays a msgbox that shows how many rows are deleted
        // MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

        // End If

        // 'if user didnt select an entire column or an entire row then following block will excecute
        // Else

        // 'if user select only a single cell the following warning will pop up and exit from the code 
        // If selectedRng.Rows.Count = 1 And selectedRng.Columns.Count = 1 Then
        // If selectedRng.Cells(1, 1).value Is Nothing Then
        // answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
        // If answer = MsgBoxResult.No Then
        // Exit Sub
        // Else
        // selectedRng.Cells(1, 1).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
        // MsgBox(1 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
        // Exit Sub
        // End If
        // Else
        // MsgBox(0 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
        // Exit Sub
        // End If
        // End If



        // 'if there is no blank cell in the selection then an exception will be thrown and 0 cells will be deleted
        // Try
        // blankCells = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks)

        // Dim arr_Blank_Cell_Address() As String = Split(blankCells.Address, ",")

        // answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
        // If answer = MsgBoxResult.No Then
        // Exit Sub
        // Else

        // For i = 0 To UBound(arr_Blank_Cell_Address)

        // If worksheet.Range(arr_Blank_Cell_Address(i)).Columns.Count = selectedRng.Columns.Count Then
        // worksheet.Range(arr_Blank_Cell_Address(i)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
        // blankRowCount += worksheet.Range(arr_Blank_Cell_Address(i)).Rows.Count
        // Else
        // Continue For
        // End If
        // Next

        // selectedRng.Cells(1, 1).select

        // 'displays a msgbox that shows how many rows are deleted
        // MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

        // End If

        // Catch ex As Exception

        // MsgBox(0 & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
        // Exit Sub

        // End Try


        // End If


        // Else

        // Dim arrRng As String() = Split(selectedRng.Address, ",")

        // Dim totalBlankRowCount As Integer = 0

        // answer = MsgBox("Confirm: Data will move up. Do you still want to proceed?", MsgBoxStyle.YesNo, "Warning!")
        // If answer = MsgBoxResult.No Then
        // Exit Sub
        // Else

        // For i = 0 To UBound(arrRng)
        // selectedRng = worksheet.Range(arrRng(i))


        // isMatchRow = Regex.IsMatch(UCase(selectedRng.Address), rowPattern)


        // 'columnPattern = "^(\$?[A-Z]+)(:)(\$?[A-Z]+)$"
        // isMatchColumn = Regex.IsMatch(UCase(selectedRng.Address), columnPattern)

        // 'checks if the user selected an entire column or not, if yes then following block will excecute
        // If isMatchColumn = True And isMatchRow = False Then

        // 'if there is no blank cell in the entire column then an exception will be thrown 
        // Try
        // blankCells = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks)
        // blankRowCount = blankCells.Cells.Count

        // blankCells.Delete(Excel.XlDeleteShiftDirection.xlShiftUp)

        // Catch ex As Exception

        // GoTo nextRange

        // End Try


        // 'if user selects an entire row as selection then following block will execute
        // ElseIf isMatchColumn = False And isMatchRow = True Then


        // blankRowCount = 0

        // For k = selectedRng.Rows.Count To 1 Step -1

        // 'checks if the entire row is empty or not
        // If excelApp.WorksheetFunction.CountA(selectedRng.Rows(k)) = 0 Then
        // blankRowCount += 1
        // selectedRng.Rows(k).delete
        // End If
        // Next



        // 'if user didnt select an entire column or an entire row then following block will excecute
        // Else
        // blankRowCount = 0
        // 'if user select only a single cell the following warning will pop up and exit from the code 
        // If selectedRng.Rows.Count = 1 And selectedRng.Columns.Count = 1 Then
        // If selectedRng.Cells(1, 1).value Is Nothing Then
        // selectedRng.Cells(1, 1).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
        // blankRowCount += 1
        // GoTo nextRange
        // Else
        // GoTo nextRange
        // End If
        // End If



        // 'if there is no blank cell in the selection then an exception will be thrown and 0 cells will be deleted
        // Try
        // blankCells = selectedRng.SpecialCells(XlCellType.xlCellTypeBlanks)

        // Dim arr_Blank_Cell_Address() As String = Split(blankCells.Address, ",")

        // For k = 0 To UBound(arr_Blank_Cell_Address)

        // If worksheet.Range(arr_Blank_Cell_Address(k)).Columns.Count = selectedRng.Columns.Count Then
        // worksheet.Range(arr_Blank_Cell_Address(k)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
        // blankRowCount += worksheet.Range(arr_Blank_Cell_Address(k)).Rows.Count
        // Else
        // GoTo nextRange
        // End If
        // Next

        // Catch ex As Exception

        // End Try

        // End If
        // nextRange:
        // totalBlankRowCount += blankRowCount
        // Next

        // End If

        // MsgBox(totalBlankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

        // End If


        // Catch ex As Exception

        // End Try

        // End Sub





        /// <summary>
    /// removes all blank rows from active worksheet
    /// </summary>

        private void Button38_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {

                Excel.Application excelApp;
                Excel.Workbook workbook;
                Excel.Worksheet worksheet;
                Range selectedRng;
                string blankRowList = "";
                int blankRowCount = 0;
                int i;
                string flag = "Empty";

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                selectedRng = (Range)excelApp.Selection;

                // use the UsedRange method to find the address of the range used in the active sheet
                // use split function to get 2nd portion of the range which is the last cell of the used range
                // Use this addrees to to find row number of last cell
                string[] lastCell;
                int lastRowNum;
                int lastColNum;

                lastCell = worksheet.UsedRange.get_Address().Split(':');
                lastRowNum = worksheet.get_Range(lastCell[1]).Row;
                lastColNum = worksheet.get_Range(lastCell[1]).Column;



                // loop through each rows of the active sheet upto the last row number
                // check if the entire row is empty or not by using the "IsEmptyRow" function



                for (i = lastRowNum; i >= 1; i -= 1)
                {
                    flag = "Empty";
                    for (int j = lastColNum; j >= 1; j -= 1)
                    {
                        if (worksheet.Cells[i, j].value is not null)
                        {

                            flag = "NotEmpty";

                        }

                    }

                    if (flag == "Empty")
                    {

                        worksheet.Cells[i, 1].entirerow.delete();

                        blankRowCount += 1;

                    }
                }

                // if no blank rows are found in a sheet then go to the "nextloop" section and skip the lines in between
                if (blankRowCount == 0)
                {
                    goto break;
                }

                worksheet.Cells[1, 1].select();

break:
                ;

                // displays a msgbox that shows how many rows are deleted
                Interaction.MsgBox(blankRowCount + " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO");
            }




            catch (Exception ex)
            {

            }




            // Dim excelApp As Excel.Application = Nothing
            // Dim workbook As Excel.Workbook = Nothing
            // Dim worksheet As Excel.Worksheet = Nothing
            // Dim selectedRng As Excel.Range = Nothing
            // Dim range As Excel.Range = Nothing
            // Dim blankRowList As String = ""
            // Dim blankRowCount As Integer = 0

            // Try
            // excelApp = Globals.ThisAddIn.Application
            // workbook = excelApp.ActiveWorkbook
            // worksheet = workbook.ActiveSheet
            // selectedRng = excelApp.Selection

            // Dim lastCell() As String
            // Dim lastRowNum As Long
            // Dim lastColNum As Long

            // lastCell = worksheet.UsedRange.Address.Split(":"c)
            // lastRowNum = worksheet.Range(lastCell(1)).Row
            // lastColNum = worksheet.Range(lastCell(1)).Column

            // range = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(lastRowNum, lastColNum))

            // MsgBox(lastRowNum)
            // MsgBox(lastColNum)



            // For i = 1 To lastRowNum
            // For j = 1 To lastColNum


            // Dim currentRow As Excel.Range = range.Rows(i)
            // 'If excelApp.WorksheetFunction.CountA(currentRow) = 0 Then
            // If Not worksheet.Cells(i, j).value IsNot Nothing AndAlso worksheet.Cells(i, j).ToString().Trim() <> "" Then
            // If i = 128 Then
            // MsgBox("y")
            // End If
            // Exit Sub
            // 'blankRowList &= "," & worksheet.Range(worksheet.Cells(i, 1), worksheet.Cells(i, 2)).Address
            // 'blankRowCount += 1
            // End If
            // 'Marshal.ReleaseComObject(currentRow)
            // Next
            // Next

            // If blankRowCount > 0 Then
            // blankRowList = Right(blankRowList, Len(blankRowList) - 1)
            // worksheet.Range(blankRowList).EntireRow.Delete()
            // worksheet.Cells(1, 1).Select()
            // MsgBox($"{blankRowCount} Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
            // End If

            // Catch ex As Exception
            // ' Handle exceptions
            // Finally
            // ' Release and cleanup COM objects
            // If Not range Is Nothing Then Marshal.ReleaseComObject(range)
            // If Not selectedRng Is Nothing Then Marshal.ReleaseComObject(selectedRng)
            // If Not worksheet Is Nothing Then Marshal.ReleaseComObject(worksheet)
            // If Not workbook Is Nothing Then Marshal.ReleaseComObject(workbook)

            // GC.Collect()
            // GC.WaitForPendingFinalizers()
            // GC.Collect()
            // GC.WaitForPendingFinalizers()
            // End Try









        }

        /// <summary>
    /// removes all blank rows from all selected worksheets
    /// </summary>

        private void Button39_Click(object sender, RibbonControlEventArgs e)
        {


            try
            {


                Excel.Application excelApp;
                Excel.Workbook workbook;
                Excel.Worksheet worksheet;
                string blankRowList;
                string confirmationMsg = "";
                int blankRowCount;
                int i = 0;
                string flag = "Empty";


                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;

                // takes the sheet names of the selected worksheets
                var selectedSheets = excelApp.ActiveWindow.SelectedSheets;
                string sheetName = "";

                // loops through each selected worksheet and concatenate all the sheet names togehter in the "sheetName" variable
                // then Right function removes the leading comma (,) from the "sheetName" variable
                foreach (Excel.Worksheet sheet in selectedSheets)
                    sheetName = sheetName + "," + sheet.Name;
                sheetName = Strings.Right(sheetName, Strings.Len(sheetName) - 1);

                // new array (arrSheetName) stores all the selected sheet names separately
                string[] arrSheetName = Strings.Split(sheetName, ",");


                // loops through each selected sheet name from the "arrSheetName" array
                // "worksheet" variable takes the sheets name from the array and makes it active worksheet
                // each time a new sheet is taken from the slected sheets, "blankRowList" resets to empty string and "blankRowCount" resets to 0
                var loopTo = Information.UBound(arrSheetName);
                for (i = 0; i <= loopTo; i++)
                {
                    blankRowList = "";
                    blankRowCount = 0;
                    worksheet = (Excel.Worksheet)workbook.Sheets[arrSheetName[i]];
                    worksheet.Activate();


                    // use the UsedRange method to find the address of the range used in the active sheet
                    // use split function to get 2nd portion of the range which is the last cell of the used range
                    // Use this addrees to to find row number of last cell
                    string[] lastCell;
                    int lastRowNum;
                    int lastColNum;

                    lastCell = worksheet.UsedRange.get_Address().Split(':');
                    lastRowNum = worksheet.get_Range(lastCell[1]).Row;
                    lastColNum = worksheet.get_Range(lastCell[1]).Column;


                    // loop through each rows of the active sheet upto the last row number
                    // check if the entire column is empty or not by using the "IsRowEmpty" function
                    for (int j = lastRowNum; j >= 1; j -= 1)
                    {
                        flag = "Empty";
                        for (int k = lastColNum; k >= 1; k -= 1)
                        {
                            if (worksheet.Cells[j, k].value is not null)
                            {

                                flag = "NotEmpty";

                            }

                        }

                        if (flag == "Empty")
                        {

                            worksheet.Cells[j, 1].entirerow.delete();

                            blankRowCount += 1;

                        }
                    }

                    // if no blank rows are found in a sheet then go to the "nextloop" section and skip the lines in between
                    if (blankRowCount == 0)
                    {
                        goto nextloop;
                    }


nextloop:
                    ;

                    // stores information about how many rows deleted from which sheet
                    confirmationMsg = confirmationMsg + blankRowCount + " Row(s) are deleted from " + arrSheetName[i] + Microsoft.VisualBasic.Constants.vbCrLf;

                }

                // finally this msgBox is shown
                Interaction.MsgBox(confirmationMsg, MsgBoxStyle.Information, "SOFTEKO");
            }

            catch (Exception ex)
            {

            }

        }

        /// <summary>
    /// removes all blank rows from all worksheets from active workbook
    /// </summary>

        private void Button40_Click(object sender, RibbonControlEventArgs e)
        {


            try
            {

                Excel.Application excelApp;
                Excel.Workbook workbook;
                Excel.Worksheet worksheet;
                string confirmationMsg = "";
                int blankRowCount;
                int i = 0;
                string flag = "Empty";

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;

                // takes the sheet names of all worksheets of the workbook
                var selectedSheets = excelApp.Sheets;
                string sheetName = "";

                // loops through each selected worksheet and concatenate all the sheet names togehter in the "sheetName" variable
                // then Right function removes the leading comma (,) from the "sheetName" variable
                foreach (Excel.Worksheet sheet in selectedSheets)
                    sheetName = sheetName + "," + sheet.Name;
                sheetName = Strings.Right(sheetName, Strings.Len(sheetName) - 1);

                // new array (arrSheetName) stores all the sheet names separately
                string[] arrSheetName = Strings.Split(sheetName, ",");


                // loops through each sheet name from the "arrSheetName" array
                // "worksheet" variable takes the sheet names from the array and makes it active worksheet
                // each time a new sheet is taken by "worksheet" variable, "blankRowList" resets to empty string and "blankRowCount" resets to 0
                var loopTo = Information.UBound(arrSheetName);
                for (i = 0; i <= loopTo; i++)
                {
                    blankRowCount = 0;
                    worksheet = (Excel.Worksheet)workbook.Sheets[arrSheetName[i]];
                    worksheet.Activate();


                    // use the UsedRange method to find the address of the range used in the active sheet
                    // use split function to get 2nd portion of the range which is the last cell of the used range
                    // Use this addrees to to find row number of last cell
                    string[] lastCell;
                    int lastRowNum;
                    int lastColNum;

                    lastCell = worksheet.UsedRange.get_Address().Split(':');
                    lastRowNum = worksheet.get_Range(lastCell[1]).Row;
                    lastColNum = worksheet.get_Range(lastCell[1]).Column;


                    // loop through each rows of the active sheet upto the last row number
                    // check if the entire column is empty or not by using the "IsRowEmpty" function
                    for (int j = lastRowNum; j >= 1; j -= 1)
                    {
                        flag = "Empty";
                        for (int k = lastColNum; k >= 1; k -= 1)
                        {
                            if (worksheet.Cells[j, k].value is not null)
                            {

                                flag = "NotEmpty";

                            }

                        }

                        if (flag == "Empty")
                        {

                            worksheet.Cells[j, 1].entirerow.delete();

                            blankRowCount += 1;

                        }
                    }


                    // if no blank rows are found in a sheet then go to the "nextloop" section and skip the lines in between
                    if (blankRowCount == 0)
                    {
                        goto nextloop;
                    }

nextloop:
                    ;

                    // stores information about how many rows deleted from which sheet
                    confirmationMsg = confirmationMsg + blankRowCount + " Row(s) are deleted from " + arrSheetName[i] + Microsoft.VisualBasic.Constants.vbCrLf;

                }

                // finally this msgBox is shown
                Interaction.MsgBox(confirmationMsg, MsgBoxStyle.Information, "SOFTEKO");
            }

            catch (Exception ex)
            {

            }


        }

        /// <summary>
    /// removes empty sheets from the active workbook
    /// </summary>

        private void Button41_Click(object sender, RibbonControlEventArgs e)
        {

            try
            {

                Excel.Application excelApp;
                Excel.Workbook workbook;
                Excel.Worksheet worksheet;
                int blankWsCount = 0;
                int i = 0;
                string flag;
                MsgBoxResult answer;
                string initialWs;

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;

                // "initialWs" variable stores the name of the worksheet, where the button event was clicked
                initialWs = Conversions.ToString(excelApp.ActiveSheet.name);

                // takes the sheet names of all worksheets of the workbook
                var selectedSheets = excelApp.Sheets;
                string sheetName = "";

                // loops through each selected worksheet and concatenate all the sheet names togehter in the "sheetName" variable
                // then Right function removes the leading comma (,) from the "sheetName" variable
                foreach (Excel.Worksheet sheet in selectedSheets)
                    sheetName = sheetName + "," + sheet.Name;
                sheetName = Strings.Right(sheetName, Strings.Len(sheetName) - 1);

                // new array (arrSheetName) stores all the sheet names separately
                string[] arrSheetName = Strings.Split(sheetName, ",");



                // this loops only counts the number of empty WS present in the active workbook
                // loops through each selected sheet name from the "arrSheetName" array
                // "worksheet" variable takes the sheets name from the array and makes it active worksheet
                var loopTo = Information.UBound(arrSheetName);
                for (i = 0; i <= loopTo; i++)
                {
                    flag = "NotEmpty";
                    worksheet = (Excel.Worksheet)workbook.Sheets[arrSheetName[i]];
                    worksheet.Activate();


                    // loop thorugh the characters of address of the used range of a worksheet
                    // check if it conrians ":". If the WS is empty, used range will be single cell and the address will not have any ":" in it
                    // so, "usedCellCount" will be 0 for an empty WS
                    int usedCellCount = 0;
                    foreach (char c in worksheet.UsedRange.get_Address())
                    {

                        if (Conversions.ToString(c) == ":")
                        {
                            usedCellCount += 1;
                        }

                    }


                    // make sure the WS is actually empty or not by checking the value of the used range (which is already a single cell)
                    // if there is no value then flag becomes "Empty"
                    if (usedCellCount == 0)
                    {
                        if (worksheet.UsedRange.get_Value() is not null)
                        {
                            flag = "NotEmpty";
                        }
                        else
                        {
                            flag = "Empty";
                        }

                    }

                    // increase the value of "blankWsCount" by 1 if an empty WS is found
                    if (flag == "Empty")
                    {

                        blankWsCount += 1;

                    }

                }

                // if no blank WS is found then display this message and exit sub
                if (blankWsCount == 0)
                {
                    Interaction.MsgBox("No empty worksheet is found.", MsgBoxStyle.Information, "SOFTEKO");
                    return;
                }

                workbook.Sheets[initialWs].activate();

                // assign the reponse of user from the msgbox to the "answer" variable
                answer = Interaction.MsgBox(blankWsCount + " empty worksheet(s) will be deleted. Please click Yes to continue.", MsgBoxStyle.YesNo, "SOFTEKO");

                if (answer == MsgBoxResult.Yes)
                {

                    // this loop deletes the empty worksheets
                    // mechanism is same as previous loop
                    var loopTo1 = Information.UBound(arrSheetName);
                    for (i = 0; i <= loopTo1; i++)
                    {
                        flag = "NotEmpty";
                        worksheet = (Excel.Worksheet)workbook.Sheets[arrSheetName[i]];
                        worksheet.Activate();

                        int usedCellCount = 0;
                        foreach (char c in worksheet.UsedRange.get_Address())
                        {

                            if (Conversions.ToString(c) == ":")
                            {
                                usedCellCount += 1;
                            }

                        }

                        if (usedCellCount == 0)
                        {
                            if (worksheet.UsedRange.get_Value() is not null)
                            {
                                flag = "NotEmpty";
                            }
                            else
                            {
                                flag = "Empty";
                            }

                        }

                        if (flag == "Empty")
                        {

                            worksheet.Delete();

                        }

                    }

                    workbook.Sheets[initialWs].activate();

                    // finally this msgBox is shown
                    Interaction.MsgBox(blankWsCount + " worksheet(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO");
                }
                else
                {
                    return;
                }
            }


            catch (Exception ex)
            {

            }

        }

        private void DropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void ComboBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void Button9_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void Button10_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void Button55_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void Button44_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void Button43_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void Button42_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void Button2_Click_1(object sender, RibbonControlEventArgs e)
        {

        }

        private void Button16_Click(object sender, RibbonControlEventArgs e)
        {

            if (GlobalModule.form_flag == false)
            {
                var MyForm18 = new Form18_CombineRanges();

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selection = (Range)excelApp.Selection;

                MyForm18.TextBox1.Text = selection.get_Address();
                MyForm18.ComboBox1.SelectedIndex = -1;
                MyForm18.ComboBox1.Text = "SOFTEKO";
                MyForm18.Show();
                MyForm18.RadioButton1.Checked = true;
                GlobalModule.form_flag = true;
            }
        }

        private void Button20_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var MyForm22 = new Form22_Merge_Duplicate_Rows();

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selection = (Range)excelApp.Selection;

                MyForm22.TextBox1.Text = selection.get_Address();
                MyForm22.ComboBox1.SelectedIndex = -1;
                MyForm22.ComboBox1.Text = "SOFTEKO";
                MyForm22.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button21_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var MyForm23 = new Form23_Merge_Duplicate_Columns();

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selection = (Range)excelApp.Selection;

                MyForm23.TextBox1.Text = selection.get_Address();
                MyForm23.ComboBox1.SelectedIndex = -1;
                MyForm23.ComboBox1.Text = "SOFTEKO";
                MyForm23.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button23_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var MyForm25 = new Form25_Split_Range();

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selection = (Range)excelApp.Selection;

                MyForm25.TextBox1.Text = selection.get_Address();
                MyForm25.ComboBox1.SelectedIndex = -1;
                MyForm25.ComboBox1.Text = "SOFTEKO";
                MyForm25.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button22_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var MyForm24 = new Form24_Split_Cells();

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                // worksheet = workbook.ActiveSheet
                MyForm24.OpenSheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selection = (Range)excelApp.Selection;

                MyForm24.TextBox1.Text = selection.get_Address();
                MyForm24.ComboBox1.SelectedIndex = -1;
                MyForm24.ComboBox1.Text = "SOFTEKO";
                MyForm24.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button45_Click(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var MyForm26 = new Form26_split_text_bycharacters();

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selection = (Range)excelApp.Selection;
                MyForm26.TB_source_range.Text = selection.get_Address();
                MyForm26.ComboBox1.SelectedIndex = -1;
                MyForm26.ComboBox1.Text = "SOFTEKO";
                MyForm26.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button46_Click(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                // Dim Source As String = "Absbsjdwd,hdwdiqd,djd"
                // Dim pattern As String = "***,*,"
                // Dim KeepSeparator As Boolean = True
                // Dim Consecutive As Boolean = True
                // Dim Before As Boolean = True

                // Dim Values() As String
                // Values = SplitText(Source, pattern, Consecutive, KeepSeparator, Before)
                // For i = LBound(Values) To UBound(Values)
                // MsgBox(Values(i))
                // Next

                var MyForm27 = new Form27_Split_text_bystrings();
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selection = (Range)excelApp.Selection;
                MyForm27.TB_source_range.Text = selection.get_Address();
                MyForm27.ComboBox1.SelectedIndex = -1;
                MyForm27.ComboBox1.Text = "SOFTEKO";
                MyForm27.Show();
                GlobalModule.form_flag = true;
            }

        }


        private void Button49_Click(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var MyForm33 = new Form33_ColorBasedDropDownList();
                MyForm33.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button54_Click(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var form = new Form12HideRanges();

                form.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button11_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var form = new Form13HideAllExceptSelectedRange();

                form.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button12_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var Myform = new Form14SpecifyScrollArea();
                Myform.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button2_Click_2(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var form = new Form29_Simple_Drop_down_List();
                form.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button9_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var form = new Form34_PictureBasedDropdownList();

                form.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button29_Click(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var form = new Form31_UpdateDynamicDropdownList();

                form.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button30_Click(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var form = new Form32_ExtendDropDownList();

                form.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button24_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                GlobalModule.settingflag1 = false;
                var form = new Form35Multi_SelectionbasedDropdown();

                form.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button25_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                GlobalModule.settingflag2 = false;
                var form = new Form37_MSDropDownCheckBox();

                form.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button26_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var form = new Form39_DropdownlistwithSearchOption();

                form.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button27_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var form = new Form41_RemoveAdavancedDropdownList();

                form.Show();
                GlobalModule.form_flag = true;
            }
        }

        private void Button17_Click(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var MyForm18 = new Form18_CombineRanges();

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selection = (Range)excelApp.Selection;

                MyForm18.TextBox1.Text = selection.get_Address();
                MyForm18.ComboBox1.SelectedIndex = -1;
                MyForm18.ComboBox1.Text = "SOFTEKO";
                MyForm18.Show();
                MyForm18.RadioButton2.Checked = true;
                GlobalModule.form_flag = true;
            }
        }

        private void Button18_Click(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var MyForm18 = new Form18_CombineRanges();

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selection = (Range)excelApp.Selection;

                MyForm18.TextBox1.Text = selection.get_Address();
                MyForm18.ComboBox1.SelectedIndex = -1;
                MyForm18.ComboBox1.Text = "SOFTEKO";
                MyForm18.Show();
                MyForm18.RadioButton3.Checked = true;
                GlobalModule.form_flag = true;
            }
        }

        private void Button47_Click(object sender, RibbonControlEventArgs e)
        {
            if (GlobalModule.form_flag == false)
            {
                var MyForm28 = new Form28_Split_text_bypattern();
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selection = (Range)excelApp.Selection;

                MyForm28.TB_source_range.Text = selection.get_Address();
                MyForm28.ComboBox1.SelectedIndex = -1;
                MyForm28.ComboBox1.Text = "SOFTEKO";
                MyForm28.Show();
                GlobalModule.form_flag = true;
            }

        }

        private void Button10_Click_1(object sender, RibbonControlEventArgs e)
        {
            // Clear Scroll Area
            try
            {
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                // checks if a scroll area is specified or not in the worksheet
                if (Form14SpecifyScrollArea.scroll_Area_Specified == true)
                {

                    // unhide all the rows and columns of the worksheet
                    worksheet.Rows.Hidden = false;
                    worksheet.Columns.Hidden = false;

                    // loop through each element of the all_hidden_Row_No list from form14, and fetch the row numbers that were hidden in the selected range
                    // hide those rows
                    for (int i = 0, loopTo = Form14SpecifyScrollArea.all_hidden_Row_No.Count - 1; i <= loopTo; i++)
                        worksheet.Rows[Form14SpecifyScrollArea.all_hidden_Row_No[i]].hidden = (object)true;

                    // loop through each element of the all_hidden_Col_No list from form14, and fetch the column numbers that were hidden in the selected range
                    // hide those columns
                    for (int i = 0, loopTo1 = Form14SpecifyScrollArea.all_hidden_Col_No.Count - 1; i <= loopTo1; i++)
                        worksheet.Columns[Form14SpecifyScrollArea.all_hidden_Col_No[i]].hidden = (object)true;


                }
            }


            catch (Exception ex)
            {

            }




        }
    }
}