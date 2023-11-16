using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace VSTO_Addins
{



    public partial class Form21FillEmtyCells
    {

        private Excel.Application _excelApp;

        private Excel.Application excelApp
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _excelApp;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _excelApp = value;
            }
        }
        private Excel.Workbook workbook;
        private Excel.Worksheet worksheet, worksheet1;
        private Excel.Worksheet outWorksheet;
        private Range inputRng;
        private int FocusedTxtBox;
        private Range selectedRange;
        private bool txtChanged = false;

        public Form21FillEmtyCells()
        {
            InitializeComponent();
        }

        [DllImport("user32")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private const uint SWP_NOMOVE = 0x2U;
        private const uint SWP_NOSIZE = 0x1U;
        private const uint SWP_NOACTIVATE = 0x10U;
        private const int HWND_TOPMOST = -1;

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_OK.PerformClick();
            }
        }

        private void RB_Linear_values_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_Linear_values.Checked == true)
            {
                ComboBox_Options.Items.Clear();
                ComboBox_Options.Items.Add("Top to Buttom");
                ComboBox_Options.Items.Add("Left to Right");
                ComboBox_Options.SelectedIndex = 0;
                txtFillValue.Enabled = false;
                L_Fill_Value.Enabled = false;
                ComboBox_Options.Enabled = true;
                L_Fill_Options.Enabled = true;
                CB_Keepformatting.Enabled = true;
            }

        }

        private void RB_Values_fromselected_range_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_Values_fromselected_range.Checked == true)
            {
                ComboBox_Options.Items.Clear();
                ComboBox_Options.Items.Add("Downwards");
                ComboBox_Options.Items.Add("Upwards");
                ComboBox_Options.Items.Add("Towards the Right");
                ComboBox_Options.Items.Add("Towards the Left");
                ComboBox_Options.SelectedIndex = 0;
                txtFillValue.Enabled = false;
                L_Fill_Value.Enabled = false;
                ComboBox_Options.Enabled = true;
                L_Fill_Options.Enabled = true;
                CB_Keepformatting.Enabled = true;
            }
        }

        private void RB_Certain_value_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_Certain_value.Checked == true)
            {
                ComboBox_Options.Items.Clear();
                ComboBox_Options.SelectedItem = "";
                txtFillValue.Enabled = true;
                L_Fill_Value.Enabled = true;
                ComboBox_Options.Enabled = false;
                L_Fill_Options.Enabled = false;
                CB_Keepformatting.Enabled = false;
            }
        }


        private void Form21FillEmtyCells_Load(object sender, EventArgs e)
        {

            try
            {

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selectedRng = (Range)excelApp.Selection;
                txtSourceRange.Text = selectedRng.get_Address();

                KeyPreview = true;
            }

            catch (Exception ex)
            {

            }


        }

        private void Selection_Click(object sender, EventArgs e)
        {


            try
            {

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                selectedRange = (Range)excelApp.Selection;
                txtSourceRange.Focus();

                Hide();
                inputRng = (Range)excelApp.InputBox("Please Select a Range", "Range Selection", selectedRange.get_Address(), Type: 8);
                Show();

                inputRng.Worksheet.Activate();

                txtSourceRange.Text = inputRng.get_Address();
                inputRng.Select();
            }


            catch (Exception ex)
            {

                txtSourceRange.Focus();

            }




        }

        private void Textbox1_TextChanged(object sender, EventArgs e)
        {

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                txtChanged = true;

                inputRng = worksheet.get_Range(txtSourceRange.Text);
                inputRng.Select();
            }


            catch (Exception ex)
            {

            }

            txtChanged = false;
            txtSourceRange.Focus();



        }



        private void txtSourceRange_GotFocus(object sender, EventArgs e)
        {
            try
            {

                FocusedTxtBox = 1;
            }


            catch (Exception ex)
            {

            }
        }
        private void Form1_Activated(object sender, EventArgs e)
        {
            try
            {

                excelApp = Globals.ThisAddIn.Application;

                excelApp.SheetSelectionChange += rngSelectionFromTxtBox;
            }

            catch (Exception ex)
            {

            }

        }

        private void rngSelectionFromTxtBox(object Sh, Range Target)
        {

            try
            {

                excelApp = Globals.ThisAddIn.Application;
                selectedRange = (Range)excelApp.Selection;
                selectedRange.Select();

                txtSourceRange.Focus();

                if (txtChanged == false)
                {

                    if (FocusedTxtBox == 1)
                    {

                        txtSourceRange.Text = selectedRange.get_Address();
                        worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                        inputRng = selectedRange;
                        txtSourceRange.Focus();

                    }

                }
            }



            catch (Exception ex)
            {


            }

        }

        public bool IsValidRng(string input)
        {

            string pattern = @"^(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$";
            return System.Text.RegularExpressions.Regex.IsMatch(input, pattern);

        }


        private void btn_OK_Click(object sender, EventArgs e)
        {

            try
            {

                string inputWsName;
                string fillValue;
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                selectedRange = (Range)excelApp.Selection;
                inputWsName = worksheet.Name;

                // checks if an empty source range is used or not
                // if it is blank then a warning msgbox will appear and give user another chance to enter source range
                // if it is not blank then it checks the used range is valid range or not by using IsValidRng() function
                // IsValidRng() function is a custom function (see line 200)
                // using invalid range will give a warning to user and give another chance to enter range correctly
                if (string.IsNullOrEmpty(txtSourceRange.Text))
                {
                    Interaction.MsgBox("Please select the Source Range.", MsgBoxStyle.Exclamation, "Error!");
                    txtSourceRange.Focus();
                    return;
                }
                else if (IsValidRng(txtSourceRange.Text.ToUpper()) == false)
                {
                    Interaction.MsgBox("Please use a valid range.", MsgBoxStyle.Exclamation, "Error!");
                    txtSourceRange.Text = "";
                    txtSourceRange.Focus();
                    return;
                }


                // stores the text value of the textbox in "temp" variable to use it later
                // store the active worksheet into "worksheet1" variable
                string temp;
                temp = txtSourceRange.Text;
                worksheet1 = inputRng.Worksheet;

                // if CB_Backup_Sheet is checked then this will copy the active sheet and reactivate the original worksheet
                // replace the text of the txtSourceRange textbox by "temp" variable
                if (CB_Backup_Sheet.Checked == true)
                {

                    workbook.ActiveSheet.Copy(After: workbook.Sheets[workbook.Sheets.Count]);
                    outWorksheet = (Excel.Worksheet)workbook.Sheets[workbook.Sheets.Count];

                    worksheet1.Activate();
                    txtSourceRange.Text = temp;

                }

                // RB_Values_fromselected_range with Downwards fill option 
                if (RB_Values_fromselected_range.Checked == true)
                {

                    if (ComboBox_Options.SelectedIndex == 0)
                    {

                        // takes all the ranges selected by user into an array named arrRng
                        string[] arrRng = Strings.Split(txtSourceRange.Text, ",");

                        // loops through each range selected by user, which is stored in arrRng array
                        for (int p = 0, loopTo = Information.UBound(arrRng); p <= loopTo; p++)
                        {
                            selectedRange = worksheet.get_Range(arrRng[p]);

                            // loops through the cells of the selected range column by column
                            for (int j = 1, loopTo1 = selectedRange.Columns.Count; j <= loopTo1; j++)
                            {

                                // checks if the first cell of the column is blank or not
                                // if so then value of fillValue var will be blank
                                // if not, fillValue will be the value of the first cell
                                if (selectedRange.Cells[1, j].value is null)
                                {
                                    fillValue = "";
                                }
                                else
                                {
                                    fillValue = Conversions.ToString(selectedRange.Cells[1, j].value);
                                }


                                for (int i = 1, loopTo2 = selectedRange.Rows.Count; i <= loopTo2; i++)
                                {

                                    // checks if the current cell is blank or not. this condition only passes from 2nd row (i=2)
                                    // if the current cell is not blank then, replace the value of fillValue by the cuurent cell value. Then it can be copied to the following cells if they are blank
                                    if (selectedRange.Cells[i, j].value is null & i > 1)
                                    {

                                        // checks if the CB_Keepformatting is checked
                                        // if so then, copy the cell of the previous row and same column (i-1,j) and paste it in current cell. This will copy both the value and format
                                        // if CB_Keepformatting is not checked then, cuurent cell's value will be the value of fillValue
                                        if (CB_Keepformatting.Checked == true)
                                        {
                                            selectedRange.Cells[i - 1, j].copy(selectedRange.Cells[i, j]);
                                        }
                                        else
                                        {
                                            selectedRange.Cells[i, j].value = fillValue;
                                        }
                                    }

                                    else
                                    {
                                        fillValue = Conversions.ToString(selectedRange.Cells[i, j].value);
                                    }

                                }
                            }

                        }
                    }



                    // RB_Values_fromselected_range with Upwards fill option 
                    else if (ComboBox_Options.SelectedIndex == 1)
                    {

                        // takes all the ranges selected by user into an array named arrRng
                        string[] arrRng = Strings.Split(txtSourceRange.Text, ",");
                        int rowCount = selectedRange.Rows.Count;

                        // loops through each range selected by user, which is stored in arrRng array
                        for (int p = 0, loopTo3 = Information.UBound(arrRng); p <= loopTo3; p++)
                        {

                            selectedRange = worksheet.get_Range(arrRng[p]);

                            // loops through the cells of the selected range column by column
                            for (int j = 1, loopTo4 = selectedRange.Columns.Count; j <= loopTo4; j++)
                            {


                                // checks if the last cell of the column is blank or not
                                // if so then value of fillValue var will be blank
                                // if not, fillValue will be the value of the last cell

                                if (selectedRange.Cells[rowCount, j].value is null)
                                {
                                    fillValue = "";
                                }
                                else
                                {
                                    fillValue = Conversions.ToString(selectedRange.Cells[rowCount, j].value);
                                }


                                for (int i = rowCount; i >= 1; i -= 1)
                                {

                                    // checks if the current cell is blank or not. this condition only passes from 2nd from last row (i < rowCount)
                                    // if the current cell is not blank then, replace the value of fillValue by the cuurent cell value. Then it can be copied to the previous cells if they are blank
                                    if (selectedRange.Cells[i, j].value is null & i < rowCount)
                                    {

                                        // checks if the CB_Keepformatting is checked
                                        // if so then, copy the cell of the next row and same column (i+1,j) and paste it in current cell. This will copy both the value and format
                                        // if CB_Keepformatting is not checked then, cuurent cell's value will be the value of fillValue
                                        if (CB_Keepformatting.Checked == true)
                                        {
                                            selectedRange.Cells[i + 1, j].copy(selectedRange.Cells[i, j]);
                                        }
                                        else
                                        {
                                            selectedRange.Cells[i, j].value = fillValue;
                                        }
                                    }

                                    else
                                    {
                                        fillValue = Conversions.ToString(selectedRange.Cells[i, j].value);
                                    }

                                }
                            }

                        }
                    }


                    // RB_Values_fromselected_range with Towards Right fill option 
                    else if (ComboBox_Options.SelectedIndex == 2)
                    {

                        // takes all the ranges selected by user into an array named arrRng
                        string[] arrRng = Strings.Split(txtSourceRange.Text, ",");

                        // loops through each range selected by user, which is stored in arrRng array
                        for (int p = 0, loopTo5 = Information.UBound(arrRng); p <= loopTo5; p++)
                        {
                            selectedRange = worksheet.get_Range(arrRng[p]);

                            // loops through the cells of the selected range row by row
                            for (int i = 1, loopTo6 = selectedRange.Rows.Count; i <= loopTo6; i++)
                            {

                                // checks if the first cell of the row is blank or not
                                // if so then value of fillValue var will be blank
                                // if not, fillValue will be the value of the first cell
                                if (selectedRange.Cells[i, 1].value is null)
                                {
                                    fillValue = "";
                                }
                                else
                                {
                                    fillValue = Conversions.ToString(selectedRange.Cells[i, 1].value);
                                }


                                for (int j = 1, loopTo7 = selectedRange.Columns.Count; j <= loopTo7; j++)
                                {

                                    // checks if the current cell is blank or not. this condition only passes from 2nd column(j > 1)
                                    // if the current cell is not blank then, replace the value of fillValue by the cuurent cell value. Then it can be copied to the previous cells if they are blank
                                    if (selectedRange.Cells[i, j].value is null & j > 1)
                                    {

                                        // checks if the CB_Keepformatting is checked
                                        // if so then, copy the cell of the previous column and same row(i,j-1) and paste it in current cell. This will copy both the value and format
                                        // if CB_Keepformatting is not checked then, cuurent cell's value will be the value of fillValue
                                        if (CB_Keepformatting.Checked == true)
                                        {
                                            selectedRange.Cells[i, j - 1].copy(selectedRange.Cells[i, j]);
                                        }
                                        else
                                        {
                                            selectedRange.Cells[i, j].value = fillValue;
                                        }
                                    }

                                    else
                                    {
                                        fillValue = Conversions.ToString(selectedRange.Cells[i, j].value);
                                    }

                                }
                            }
                        }
                    }



                    // RB_Values_fromselected_range with Towards Left fill option 
                    else if (ComboBox_Options.SelectedIndex == 3)
                    {


                        // takes all the ranges selected by user into an array named arrRng
                        string[] arrRng = Strings.Split(txtSourceRange.Text, ",");
                        int colCount = selectedRange.Columns.Count;

                        // loops through each range selected by user, which is stored in arrRng array
                        for (int p = 0, loopTo8 = Information.UBound(arrRng); p <= loopTo8; p++)
                        {

                            selectedRange = worksheet.get_Range(arrRng[p]);

                            // loops through the cells of the selected range row by row
                            for (int i = 1, loopTo9 = selectedRange.Rows.Count; i <= loopTo9; i++)
                            {

                                // checks if the last cell of the row is blank or not
                                // if so then value of fillValue var will be blank
                                // if not, fillValue will be the value of the last cell
                                if (selectedRange.Cells[i, colCount].value is null)
                                {
                                    fillValue = "";
                                }
                                else
                                {
                                    fillValue = Conversions.ToString(selectedRange.Cells[i, colCount].value);
                                }


                                for (int j = colCount; j >= 1; j -= 1)
                                {

                                    // checks if the current cell is blank or not. this condition only passes from 2nd last column(j < colCount)
                                    // if the current cell is not blank then, replace the value of fillValue by the cuurent cell value. Then it can be copied to the previous cells if they are blank
                                    if (selectedRange.Cells[i, j].value is null & j < colCount)
                                    {

                                        // checks if the CB_Keepformatting is checked
                                        // if so then, copy the cell of the next column and same row(i,j+1) and paste it in current cell. This will copy both the value and format
                                        // if CB_Keepformatting is not checked then, cuurent cell's value will be the value of fillValue
                                        if (CB_Keepformatting.Checked == true)
                                        {
                                            selectedRange.Cells[i, j + 1].copy(selectedRange.Cells[i, j]);
                                        }
                                        else
                                        {
                                            selectedRange.Cells[i, j].value = fillValue;
                                        }
                                    }

                                    else
                                    {
                                        fillValue = Conversions.ToString(selectedRange.Cells[i, j].value);
                                    }

                                }
                            }
                        }


                    }
                }




                else if (RB_Linear_values.Checked == true)
                {
                    double startValue, endValue, steps;
                    Range startCell;

                    // RB_Linear_values selected with Top to Bottom fill option 
                    if (ComboBox_Options.SelectedIndex == 0)
                    {


                        // takes all the ranges selected by user into an array named arrRng
                        string[] arrRng = Strings.Split(txtSourceRange.Text, ",");
                        string tempRng;


                        // loops through each range selected by user, which is stored in arrRng array
                        for (int p = 0, loopTo10 = Information.UBound(arrRng); p <= loopTo10; p++)
                        {

                            selectedRange = worksheet.get_Range(arrRng[p]);

                            tempRng = arrRng[p];
                            // loops through the each cells row by row
                            for (int j = 1, loopTo11 = selectedRange.Columns.Count; j <= loopTo11; j++)
                            {

                                startValue = 0d;
                                endValue = 0d;
                                startCell = null;

                                for (int i = 1, loopTo12 = selectedRange.Rows.Count; i <= loopTo12; i++)
                                {

                                    // checks if the current cell is blank or not and makes sure that it is numeric value
                                    if (selectedRange.Cells[i, j].value is not null && Information.IsNumeric(selectedRange.Cells[i, j].value))
                                    {

                                        // for the first non empty cell of each column the startCell will be nothing and enter the first If Else block
                                        // for the following non empty cells of the column the next If Else block will be executed
                                        if (startCell is null)
                                        {
                                            startCell = (Range)selectedRange.Cells[i, j];
                                            startValue = Conversions.ToDouble(selectedRange.Cells[i, j].value);
                                        }
                                        else
                                        {

                                            endValue = Conversions.ToDouble(selectedRange.Cells[i, j].value);
                                            steps = Conversions.ToDouble(Operators.DivideObject(endValue - startValue, Operators.SubtractObject(selectedRange.Cells[i, j].Row, startCell.Row)));

                                            // fill the empty cells in between, linearly
                                            // copy formatting if CB_Keepformatting is checked, otherwise only value will be visible in the empty cells
                                            if (CB_Keepformatting.Checked == true)
                                            {
                                                for (int k = 1, loopTo13 = Conversions.ToInteger(Operators.SubtractObject(Operators.SubtractObject(selectedRange.Cells[i, j].Row, startCell.Row), 1)); k <= loopTo13; k++)
                                                {
                                                    startCell.get_Offset(k, 0).set_Value(value: Operators.AddObject(startValue, Operators.MultiplyObject(k, steps)));
                                                    startCell.Copy();
                                                    startCell.get_Offset(k, 0).PasteSpecial(XlPasteType.xlPasteFormats);
                                                }
                                                selectedRange = worksheet.get_Range(tempRng);
                                            }
                                            else
                                            {
                                                for (int k = 1, loopTo14 = Conversions.ToInteger(Operators.SubtractObject(Operators.SubtractObject(selectedRange.Cells[i, j].Row, startCell.Row), 1)); k <= loopTo14; k++)
                                                    startCell.get_Offset(k, 0).set_Value(value: Operators.AddObject(startValue, Operators.MultiplyObject(k, steps)));
                                            }

                                            // reset the value for next iteration
                                            // this block of code converts the endValue of the current iteration to the startValue for next iteration
                                            startCell = (Range)selectedRange.Cells[i, j];
                                            startValue = Conversions.ToDouble(selectedRange.Cells[i, j].value);
                                        }
                                    }

                                }
                            }
                        }
                    }



                    // RB_Linear_values selected with Left to Right fill option 
                    else if (ComboBox_Options.SelectedIndex == 1)
                    {

                        // takes all the ranges selected by user into an array named arrRng
                        string[] arrRng = Strings.Split(txtSourceRange.Text, ",");
                        string tempRng;

                        // loops through each range selected by user, which is stored in arrRng array
                        for (int p = 0, loopTo15 = Information.UBound(arrRng); p <= loopTo15; p++)
                        {

                            selectedRange = worksheet.get_Range(arrRng[p]);

                            tempRng = arrRng[p];
                            // loops through the each cells row by row
                            for (int i = 1, loopTo16 = selectedRange.Rows.Count; i <= loopTo16; i++)
                            {

                                startValue = 0d;
                                endValue = 0d;
                                startCell = null;

                                for (int j = 1, loopTo17 = selectedRange.Columns.Count; j <= loopTo17; j++)
                                {


                                    // checks if the current cell is blank or not and makes sure that it is numeric value
                                    if (selectedRange.Cells[i, j].value is not null && Information.IsNumeric(selectedRange.Cells[i, j].value))
                                    {

                                        // for the first non empty cell of each row the startCell will be nothing and enter the first If Else block
                                        // for the following non empty cells of the column the next If Else block will be executed
                                        if (startCell is null)
                                        {
                                            startCell = (Range)selectedRange.Cells[i, j];
                                            startValue = Conversions.ToDouble(selectedRange.Cells[i, j].value);
                                        }
                                        else
                                        {

                                            endValue = Conversions.ToDouble(selectedRange.Cells[i, j].value);
                                            steps = Conversions.ToDouble(Operators.DivideObject(endValue - startValue, Operators.SubtractObject(selectedRange.Cells[i, j].Column, startCell.Column)));


                                            // fill the empty cells in between, linearly
                                            // copy formatting if CB_Keepformatting is checked, otherwise only value will be visible in the empty cells
                                            if (CB_Keepformatting.Checked == true)
                                            {
                                                for (int k = 1, loopTo18 = Conversions.ToInteger(Operators.SubtractObject(Operators.SubtractObject(selectedRange.Cells[i, j].Column, startCell.Column), 1)); k <= loopTo18; k++)
                                                {
                                                    startCell.get_Offset(0, k).set_Value(value: Operators.AddObject(startValue, Operators.MultiplyObject(k, steps)));
                                                    startCell.Copy();
                                                    startCell.get_Offset(0, k).PasteSpecial(XlPasteType.xlPasteFormats);
                                                }
                                                selectedRange = worksheet.get_Range(tempRng);
                                            }
                                            else
                                            {
                                                for (int k = 1, loopTo19 = Conversions.ToInteger(Operators.SubtractObject(Operators.SubtractObject(selectedRange.Cells[i, j].Column, startCell.Column), 1)); k <= loopTo19; k++)
                                                    startCell.get_Offset(0, k).set_Value(value: Operators.AddObject(startValue, Operators.MultiplyObject(k, steps)));
                                            }

                                            // reset the value for next iteration
                                            // this block of code converts the endValue of the current iteration to the startValue for next iteration
                                            startCell = (Range)selectedRange.Cells[i, j];
                                            startValue = Conversions.ToDouble(selectedRange.Cells[i, j].value);
                                        }
                                    }

                                }
                            }
                        }


                    }
                }



                // RB_Certain_value selected
                else if (RB_Certain_value.Checked == true)
                {

                    // checks if the an empty Fill Value is used or not
                    // if so then, a warning msgbox will pop up and give user another chance to enter Fill Value
                    if (string.IsNullOrEmpty(txtFillValue.Text))
                    {
                        Interaction.MsgBox("Please enter a Fill Value.", MsgBoxStyle.Exclamation, "Error!");
                        txtFillValue.Focus();
                        return;
                    }

                    // takes all the ranges selected by user into an array named arrRng
                    string[] arrRng = Strings.Split(txtSourceRange.Text, ",");

                    // loops through each range selected by user, which is stored in arrRng array
                    for (int p = 0, loopTo20 = Information.UBound(arrRng); p <= loopTo20; p++)
                    {

                        selectedRange = worksheet.get_Range(arrRng[p]);


                        // loops through each cell of the selected range
                        for (int i = 1, loopTo21 = selectedRange.Rows.Count; i <= loopTo21; i++)
                        {
                            for (int j = 1, loopTo22 = selectedRange.Columns.Count; j <= loopTo22; j++)
                            {

                                // checks if the current cell is blank or not
                                // if so then, its cell value will be the specified Fill Value
                                if (selectedRange.Cells[i, j].value is null)
                                {
                                    selectedRange.Cells[i, j].value = txtFillValue.Text;
                                }
                            }
                        }
                    }


                }



                Dispose();
            }


            catch (Exception ex)
            {

            }

        }


        private void btn_Cancel_Click(object sender, EventArgs e)
        {

            Dispose();

        }

        private void Form21FillEmtyCells_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form21FillEmtyCells_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form21FillEmtyCells_Shown(object sender, EventArgs e)
        {
            Focus();
            BringToFront();
            Activate();
            BeginInvoke(new System.Action(() =>
                {
                    txtSourceRange.Text = inputRng.get_Address();
                    SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                }));
        }
    }
}