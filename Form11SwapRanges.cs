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

    public partial class Form11SwapRanges
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
        private Excel.Worksheet worksheet;
        private Excel.Worksheet worksheet1, worksheet2;
        private Excel.Worksheet outWorksheet;
        private Range firstInputRng;
        private Range secondInputRng;
        private int FocusedTxtBox;
        private Range selectedRange;
        private int firstRngRows, firstRngCols;
        private Range tempRng;
        private string rng1_Address, rng2_Address, initialWsName;
        private bool changeState = false;
        private bool txtChanged = false;

        public Form11SwapRanges()
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
                btnOK.PerformClick();
            }
        }

        private void Form11SwapRanges_Load(object sender, EventArgs e)
        {

            excelApp = Globals.ThisAddIn.Application;
            workbook = excelApp.ActiveWorkbook;
            worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            Range selectedRng = (Range)excelApp.Selection;
            txtSourceRange1.Text = selectedRng.get_Address();
            txtSourceRange1.Focus();

            initialWsName = worksheet.Name;

            radBtnValues.Checked = true;

            KeyPreview = true;


        }

        private void txtSourceRange1_TextChanged(object sender, EventArgs e)
        {


            try
            {

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;


                txtChanged = true;
                firstInputRng = worksheet.get_Range(txtSourceRange1.Text);


                lblSourceRng1.Text = "1st Source Range (" + firstInputRng.Rows.Count + " rows x " + firstInputRng.Columns.Count + " columns)";

                firstInputRng.Select();


                firstRngRows = worksheet.get_Range(txtSourceRange1.Text).Rows.Count;
                firstRngCols = worksheet.get_Range(txtSourceRange1.Text).Columns.Count;


                if ((firstInputRng.Worksheet.Name ?? "") != (initialWsName ?? ""))
                {

                    txtSourceRange1.Text = firstInputRng.Worksheet.Name + "!" + firstInputRng.get_Address();

                }
            }


            catch (Exception ex)
            {

            }

            txtChanged = false;

            txtSourceRange1.Focus();
        }

        private void txtSourceRange2_TextChanged(object sender, EventArgs e)
        {

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                changeState = true;

                txtChanged = true;
                secondInputRng = worksheet.get_Range(txtSourceRange2.Text);

                lblSourceRng2.Text = "2nd Source Range (" + secondInputRng.Rows.Count + " rows x " + secondInputRng.Columns.Count + " columns)";

                secondInputRng.Select();


                if ((secondInputRng.Worksheet.Name ?? "") != (initialWsName ?? ""))
                {

                    txtSourceRange2.Text = secondInputRng.Worksheet.Name + "!" + secondInputRng.get_Address();

                }
            }



            catch (Exception ex)
            {

            }

            txtChanged = false;
            txtSourceRange2.Focus();


        }

        private void rngSelection1_Click(object sender, EventArgs e)
        {

            try
            {

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                selectedRange = (Range)excelApp.Selection;
                txtSourceRange1.Focus();

                Hide();
                firstInputRng = (Range)excelApp.InputBox("Please Select the First Range", "First Range Selection", selectedRange.get_Address(), Type: 8);
                Show();



                firstInputRng.Worksheet.Activate();


                // txtSourceRange1.Text = firstInputRng.Worksheet.Name & "!" & firstInputRng.Address
                txtSourceRange1.Text = firstInputRng.get_Address();

                firstInputRng.Select();

                txtSourceRange1.Focus();
            }



            catch (Exception ex)
            {

                txtSourceRange1.Focus();

            }


        }

        private void rngSelection2_Click(object sender, EventArgs e)
        {

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                selectedRange = (Range)excelApp.Selection;
                txtSourceRange2.Focus();

                Hide();
                secondInputRng = (Range)excelApp.InputBox("Please Select the Second Range", "Second Range Selection", selectedRange.get_Address(), Type: 8);
                Show();

                secondInputRng.Worksheet.Activate();


                txtSourceRange2.Text = secondInputRng.Worksheet.Name + "!" + secondInputRng.get_Address();

                secondInputRng.Select();
                txtSourceRange2.Focus();
            }




            catch (Exception ex)
            {

                txtSourceRange2.Focus();

            }


        }



        private void AutoSelection1_Click(object sender, EventArgs e)
        {

            try
            {

                // excelApp = Globals.ThisAddIn.Application
                // workbook = excelApp.ActiveWorkbook
                // worksheet = workbook.ActiveSheet
                // selectedRange = excelApp.Selection
                // selectedRange = selectedRange.Cells(1, 1)
                // selectedRange.Select()

                // Dim topLeft, bottomRight As String



                // If selectedRange.Offset(0, -1).Value = Nothing And selectedRange.Offset(0, 1).Value = Nothing And selectedRange.Offset(-1, 0).Value = Nothing Then
                // topLeft = selectedRange.Address
                // bottomRight = worksheet.Range(topLeft).End(XlDirection.xlDown).Address
                // selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

                // ElseIf selectedRange.Offset(-1, 0).Value = Nothing And selectedRange.Offset(1, 0).Value = Nothing And selectedRange.Offset(0, -1).Value = Nothing Then

                // topLeft = selectedRange.Address
                // bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                // selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

                // ElseIf selectedRange.Offset(0, -1).Value = Nothing And selectedRange.Offset(-1, 0).Value = Nothing Then
                // bottomRight = selectedRange.End(XlDirection.xlToRight).Address
                // bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                // selectedRange = worksheet.Range(selectedRange, worksheet.Range(bottomRight))

                // ElseIf selectedRange.Offset(0, -1).Value = Nothing And selectedRange.Offset(0, 1).Value = Nothing Then

                // topLeft = selectedRange.End(XlDirection.xlUp).Address
                // bottomRight = worksheet.Range(topLeft).End(XlDirection.xlDown).Address
                // selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

                // ElseIf selectedRange.Offset(-1, 0).Value = Nothing And selectedRange.Offset(1, 0).Value = Nothing Then
                // topLeft = selectedRange.End(XlDirection.xlToLeft).Address
                // bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                // selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

                // ElseIf selectedRange.Offset(0, -1).Value = Nothing Then
                // topLeft = selectedRange.End(XlDirection.xlUp).Address
                // bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                // bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address
                // selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))


                // ElseIf selectedRange.Offset(-1, 0).Value = Nothing Then

                // topLeft = selectedRange.End(XlDirection.xlToLeft).Address
                // bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                // bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address
                // selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))



                // Else
                // topLeft = selectedRange.End(XlDirection.xlToLeft).Address
                // topLeft = worksheet.Range(topLeft).End(XlDirection.xlUp).Address
                // bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                // bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                // selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))


                // End If

                // selectedRange.Select()

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                selectedRange = (Range)excelApp.Selection;

                var activeRange = excelApp.ActiveCell;

                int startRow = activeRange.Row;
                int startColumn = activeRange.Column;
                int endRow = activeRange.Row;
                int endColumn = activeRange.Column;

                // Find the upper boundary
                while (startRow > 1 && !(worksheet.Cells[startRow - 1, startColumn].Value == null))
                    startRow -= 1;

                // Find the lower boundary
                while (!(worksheet.Cells[endRow + 1, endColumn].Value == null))
                    endRow += 1;

                // Find the left boundary
                while (startColumn > 1 && !(worksheet.Cells[startRow, startColumn - 1].Value == null))
                    startColumn -= 1;

                // Find the right boundary
                while (!(worksheet.Cells[endRow, endColumn + 1].Value == null))
                    endColumn += 1;

                // Select the determined range
                worksheet.get_Range(worksheet.Cells[startRow, startColumn], worksheet.Cells[endRow, endColumn]).Select();





                firstInputRng = selectedRange;
                txtSourceRange1.Text = firstInputRng.get_Address();

                firstRngRows = selectedRange.Rows.Count;
                firstRngCols = selectedRange.Columns.Count;
            }



            catch (Exception ex)
            {

            }


        }

        private void AutoSelection2_Click(object sender, EventArgs e)
        {


            Range firstCell;

            excelApp = Globals.ThisAddIn.Application;
            workbook = excelApp.ActiveWorkbook;
            worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            selectedRange = (Range)excelApp.Selection;
            selectedRange.Select();

            string bottomRight;
            firstCell = (Range)selectedRange.Cells[1, 1];

            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(selectedRange.Cells[1, 1].Offset((object)1, (object)0).Value, null, false)))
            {

                for (int i = 0, loopTo = firstRngCols - 1; i <= loopTo; i++)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(selectedRange.Cells[1, 1].offset((object)0, i).value, null, false)))
                    {
                        selectedRange = worksheet.get_Range(selectedRange.Cells[1, 1], selectedRange.Cells[1, 1].Offset((object)0, i));
                    }
                    selectedRange.Select();
                }
            }

            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(selectedRange.Cells[1, 1].Offset((object)0, (object)1).Value, null, false)))
            {
                for (int i = 0, loopTo1 = firstRngRows - 1; i <= loopTo1; i++)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(selectedRange.Cells[1, 1].offset(i, (object)0).value, null, false)))
                    {
                        selectedRange = worksheet.get_Range(selectedRange.Cells[1, 1], selectedRange.Cells[1, 1].Offset(i, (object)0));
                    }
                    selectedRange.Select();
                }
            }

            else
            {

                bottomRight = firstCell.get_End(XlDirection.xlToRight).get_Address();
                bottomRight = worksheet.get_Range(bottomRight).get_End(XlDirection.xlDown).get_Address();

                selectedRange = worksheet.get_Range(firstCell, worksheet.get_Range(bottomRight));

                if (selectedRange.Rows.Count == 1 & selectedRange.Columns.Count >= firstRngCols)
                {
                    selectedRange = worksheet.get_Range(selectedRange.Cells[1, 1], selectedRange.Cells[1, 1].Offset((object)0, (object)(firstRngCols - 1)));
                    selectedRange.Select();
                }

                else if (selectedRange.Rows.Count == 1 & selectedRange.Columns.Count < firstRngCols)
                {
                    selectedRange = worksheet.get_Range(selectedRange.Cells[1, 1], selectedRange.Cells[1, 1].Offset((object)0, (object)(selectedRange.Columns.Count - 1)));
                    selectedRange.Select();
                }

                else if (selectedRange.Columns.Count == 1 & selectedRange.Rows.Count >= firstRngRows)
                {
                    selectedRange = worksheet.get_Range(selectedRange.Cells[1, 1], selectedRange.Cells[1, 1].Offset((object)(firstRngRows - 1), (object)0));
                    selectedRange.Select();
                }

                else if (selectedRange.Columns.Count == 1 & selectedRange.Rows.Count < firstRngRows)
                {
                    selectedRange = worksheet.get_Range(selectedRange.Cells[1, 1], selectedRange.Cells[1, 1].Offset((object)(selectedRange.Rows.Count - 1), (object)0));
                    selectedRange.Select();
                }


                else
                {
                    bottomRight = firstCell.get_End(XlDirection.xlToRight).get_Address();
                    bottomRight = worksheet.get_Range(bottomRight).get_End(XlDirection.xlDown).get_Address();

                    selectedRange = worksheet.get_Range(firstCell, worksheet.get_Range(bottomRight));

                    if (selectedRange.Rows.Count == firstRngRows & selectedRange.Columns.Count == firstRngCols)
                    {
                        firstCell = (Range)selectedRange.Cells[1, 1];
                        selectedRange = worksheet.get_Range(firstCell.get_Offset(0, 0), firstCell.get_Offset(firstRngRows - 1, firstRngCols - 1));
                        selectedRange.Select();
                    }

                    else if (selectedRange.Rows.Count == firstRngRows & selectedRange.Columns.Count > firstRngCols)
                    {
                        firstCell = (Range)selectedRange.Cells[1, 1];
                        selectedRange = worksheet.get_Range(firstCell.get_Offset(0, 0), firstCell.get_Offset(firstRngRows - 1, firstRngCols - 1));
                        selectedRange.Select();
                    }

                    else if (selectedRange.Rows.Count == firstRngRows & selectedRange.Columns.Count < firstRngCols)
                    {
                        firstCell = (Range)selectedRange.Cells[1, 1];
                        bottomRight = firstCell.get_End(XlDirection.xlToRight).get_Address();
                        bottomRight = worksheet.get_Range(bottomRight).get_End(XlDirection.xlDown).get_Address();

                        selectedRange = worksheet.get_Range(firstCell.get_Offset(0, 0), worksheet.get_Range(bottomRight));
                        selectedRange.Select();
                    }

                    else if (selectedRange.Rows.Count > firstRngRows & selectedRange.Columns.Count == firstRngCols)
                    {
                        firstCell = (Range)selectedRange.Cells[1, 1];
                        selectedRange = worksheet.get_Range(firstCell.get_Offset(0, 0), firstCell.get_Offset(firstRngRows - 1, firstRngCols - 1));
                        selectedRange.Select();
                    }

                    else if (selectedRange.Rows.Count > firstRngRows & selectedRange.Columns.Count > firstRngCols)
                    {
                        firstCell = (Range)selectedRange.Cells[1, 1];
                        selectedRange = worksheet.get_Range(firstCell.get_Offset(0, 0), firstCell.get_Offset(firstRngRows - 1, firstRngCols - 1));
                        selectedRange.Select();
                    }

                    else if (selectedRange.Rows.Count > firstRngRows & selectedRange.Columns.Count < firstRngCols)
                    {
                        firstCell = (Range)selectedRange.Cells[1, 1];
                        bottomRight = firstCell.get_End(XlDirection.xlToRight).get_Address();
                        bottomRight = worksheet.get_Range(bottomRight).get_Offset(firstRngRows - 1, 0).get_Address();

                        selectedRange = worksheet.get_Range(firstCell.get_Offset(0, 0), worksheet.get_Range(bottomRight));
                        selectedRange.Select();
                    }

                    else if (selectedRange.Rows.Count < firstRngRows & selectedRange.Columns.Count == firstRngCols)
                    {
                        firstCell = (Range)selectedRange.Cells[1, 1];
                        bottomRight = firstCell.get_End(XlDirection.xlToRight).get_Address();
                        bottomRight = worksheet.get_Range(bottomRight).get_End(XlDirection.xlDown).get_Address();

                        selectedRange = worksheet.get_Range(firstCell.get_Offset(0, 0), worksheet.get_Range(bottomRight));
                        selectedRange.Select();
                    }
                    else if (selectedRange.Rows.Count < firstRngRows & selectedRange.Columns.Count > firstRngCols)
                    {

                        firstCell = (Range)selectedRange.Cells[1, 1];
                        bottomRight = firstCell.get_Offset(0, firstRngCols - 1).get_Address();
                        bottomRight = worksheet.get_Range(bottomRight).get_End(XlDirection.xlDown).get_Address();

                        selectedRange = worksheet.get_Range(firstCell.get_Offset(0, 0), worksheet.get_Range(bottomRight));
                        selectedRange.Select();
                    }


                    else if (selectedRange.Rows.Count < firstRngRows & selectedRange.Columns.Count < firstRngCols)
                    {
                        firstCell = (Range)selectedRange.Cells[1, 1];
                        bottomRight = firstCell.get_End(XlDirection.xlToRight).get_Address();
                        bottomRight = worksheet.get_Range(bottomRight).get_End(XlDirection.xlDown).get_Address();

                        selectedRange = worksheet.get_Range(firstCell.get_Offset(0, 0), worksheet.get_Range(bottomRight));
                        selectedRange.Select();

                    }
                }

            }

            secondInputRng = selectedRange;
            txtSourceRange2.Text = secondInputRng.get_Address();


        }


        private void txtSourceRange1_GotFocus(object sender, EventArgs e)
        {
            try
            {

                FocusedTxtBox = 1;
            }


            catch (Exception ex)
            {

            }
        }
        private void txtSourceRange2_GotFocus(object sender, EventArgs e)
        {
            try
            {

                FocusedTxtBox = 2;
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
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                selectedRange = (Range)excelApp.Selection;
                selectedRange.Select();


                if (txtChanged == false)
                {


                    if (FocusedTxtBox == 1)
                    {
                        txtSourceRange1.Text = selectedRange.get_Address();
                        txtSourceRange1.Focus();
                    }

                    else if (FocusedTxtBox == 2)
                    {
                        txtSourceRange2.Text = selectedRange.get_Address();
                    }

                }
            }


            catch (Exception ex)
            {

            }




        }

        private void btnCancel_Click(object sender, EventArgs e)
        {

            Dispose();

        }
        public bool IsValidRng(string input)
        {
            // "^(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$"

            string pattern = @"^(.*!)?(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$";
            return System.Text.RegularExpressions.Regex.IsMatch(input, pattern);

        }

        private void btnOK_Click(object sender, EventArgs e)
        {



            try
            {

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                selectedRange = (Range)excelApp.Selection;

                if (string.IsNullOrEmpty(txtSourceRange1.Text) & string.IsNullOrEmpty(txtSourceRange2.Text))
                {

                    Interaction.MsgBox("Please select the first and the second range.", MsgBoxStyle.Exclamation, "Error!");
                    txtSourceRange1.Focus();
                    return;
                }
                else if (string.IsNullOrEmpty(txtSourceRange1.Text) & !string.IsNullOrEmpty(txtSourceRange2.Text))
                {

                    if (IsValidRng(txtSourceRange2.Text.ToUpper()) == true)
                    {
                        Interaction.MsgBox("Please select the first range.", MsgBoxStyle.Exclamation, "Error!");
                        txtSourceRange1.Focus();
                        return;
                    }
                    else
                    {
                        Interaction.MsgBox("Please use a valid range in the 2nd Source Range.", MsgBoxStyle.Exclamation, "Error!");
                        txtSourceRange2.Text = "";
                        txtSourceRange2.Focus();
                        return;
                    }
                }

                else if (string.IsNullOrEmpty(txtSourceRange2.Text) & !string.IsNullOrEmpty(txtSourceRange1.Text))
                {
                    if (IsValidRng(txtSourceRange1.Text.ToUpper()) == true)
                    {
                        Interaction.MsgBox("Please select the second range.", MsgBoxStyle.Exclamation, "Error!");
                        txtSourceRange2.Focus();
                        return;
                    }
                    else
                    {
                        Interaction.MsgBox("Please use a valid range in the 1st Source Range.", MsgBoxStyle.Exclamation, "Error!");
                        txtSourceRange1.Text = "";
                        txtSourceRange1.Focus();
                        return;
                    }
                }

                else if (!string.IsNullOrEmpty(txtSourceRange1.Text) & !string.IsNullOrEmpty(txtSourceRange2.Text))
                {
                    if (IsValidRng(txtSourceRange1.Text.ToUpper()) == false & IsValidRng(txtSourceRange2.Text.ToUpper()) == true)
                    {
                        Interaction.MsgBox("Please use a valid range in the 1st Source Range.", MsgBoxStyle.Exclamation, "Error!");
                        txtSourceRange1.Text = "";
                        txtSourceRange1.Focus();
                        return;
                    }

                    else if (IsValidRng(txtSourceRange1.Text.ToUpper()) == true & IsValidRng(txtSourceRange2.Text.ToUpper()) == false)
                    {
                        Interaction.MsgBox("Please use a valid range in the 2nd Source Range.", MsgBoxStyle.Exclamation, "Error!");
                        txtSourceRange2.Text = "";
                        txtSourceRange2.Focus();
                        return;
                    }
                    else if (IsValidRng(txtSourceRange1.Text.ToUpper()) == false & IsValidRng(txtSourceRange2.Text.ToUpper()) == false)
                    {
                        Interaction.MsgBox("Please use a valid range in the Source Ranges.", MsgBoxStyle.Exclamation, "Error!");
                        txtSourceRange1.Text = "";
                        txtSourceRange2.Text = "";
                        txtSourceRange1.Focus();
                        return;

                    }
                }

                if (firstInputRng.Rows.Count != secondInputRng.Rows.Count & firstInputRng.Columns.Count != secondInputRng.Columns.Count)
                {

                    Interaction.MsgBox("You must use same number of rows and columns in both ranges.", Title: "Warning!");
                    txtSourceRange2.Focus();
                    return;
                }

                else if (firstInputRng.Rows.Count != secondInputRng.Rows.Count & firstInputRng.Columns.Count == secondInputRng.Columns.Count)
                {
                    Interaction.MsgBox("Please match the source range row size.", Title: "Warning!");
                    txtSourceRange2.Focus();
                    // Me.Dispose()
                    return;
                }
                else if (firstInputRng.Rows.Count == secondInputRng.Rows.Count & firstInputRng.Columns.Count != secondInputRng.Columns.Count)
                {
                    Interaction.MsgBox("Please match the source range column size.", Title: "Warning!");
                    txtSourceRange2.Focus();
                    return;

                }

                worksheet1 = (Excel.Worksheet)workbook.Sheets[firstInputRng.Worksheet.Name];
                worksheet2 = (Excel.Worksheet)workbook.Sheets[secondInputRng.Worksheet.Name];

                // firstInputRng = worksheet.Range(txtSourceRange1.Text)
                // secondInputRng = worksheet.Range(txtSourceRange2.Text)

                // MsgBox(worksheet1.Name)
                // MsgBox(worksheet2.Name)



                object temp;
                tempRng = worksheet1.get_Range("A10000");
                tempRng = worksheet1.get_Range(tempRng.Cells[1, 1].offset((object)0, (object)0), tempRng.Cells[1, 1].offset((object)(firstInputRng.Rows.Count - 1), (object)(firstInputRng.Columns.Count - 1)));



                if (CB_CopyWs.Checked == true)
                {

                    workbook.ActiveSheet.Copy(After: workbook.Sheets[workbook.Sheets.Count]);
                    outWorksheet = (Excel.Worksheet)workbook.Sheets[workbook.Sheets.Count];


                    worksheet = (Excel.Worksheet)workbook.Sheets[firstInputRng.Worksheet.Name];
                    worksheet.Activate();


                }

                if (radBtnValues.Checked == true)
                {
                    if (CB_KeepFormatting.Checked == true)
                    {

                        temp = firstInputRng.get_Value();
                        firstInputRng.set_Value(value: secondInputRng.get_Value());
                        secondInputRng.set_Value(value: temp);

                        for (int i = 0, loopTo = firstInputRng.Rows.Count - 1; i <= loopTo; i++)
                        {
                            for (int j = 0, loopTo1 = firstInputRng.Columns.Count - 1; j <= loopTo1; j++)
                            {


                                copyCell((Range)tempRng.Cells[1, 1], i, j, (Range)worksheet1.get_Range(firstInputRng.get_Address()).Cells[1, 1], i, j);
                                copyCell((Range)worksheet1.get_Range(firstInputRng.get_Address()).Cells[1, 1], i, j, (Range)worksheet2.get_Range(secondInputRng.get_Address()).Cells[1, 1], i, j);
                                copyCell((Range)worksheet2.get_Range(secondInputRng.get_Address()).Cells[1, 1], i, j, (Range)tempRng.Cells[1, 1], i, j);

                            }
                        }
                        tempRng.Delete();
                    }


                    else
                    {
                        firstInputRng.ClearFormats();
                        secondInputRng.ClearFormats();

                        temp = firstInputRng.get_Value();
                        firstInputRng.set_Value(value: secondInputRng.get_Value());
                        secondInputRng.set_Value(value: temp);

                    }
                    worksheet1.Activate();
                    firstInputRng.Select();
                }



                else if (radBtnKeepRef.Checked == true)
                {
                    string modifiedFormula1, modifiedFormula2;
                    if (CB_KeepFormatting.Checked == true)
                    {
                        for (int i = 0, loopTo2 = firstInputRng.Rows.Count - 1; i <= loopTo2; i++)
                        {
                            for (int j = 0, loopTo3 = firstInputRng.Columns.Count - 1; j <= loopTo3; j++)
                            {

                                copyCell((Range)tempRng.Cells[1, 1], i, j, (Range)worksheet1.get_Range(firstInputRng.get_Address()).Cells[1, 1], i, j);
                                copyCell((Range)worksheet1.get_Range(firstInputRng.get_Address()).Cells[1, 1], i, j, (Range)worksheet2.get_Range(secondInputRng.get_Address()).Cells[1, 1], i, j);
                                copyCell((Range)worksheet2.get_Range(secondInputRng.get_Address()).Cells[1, 1], i, j, (Range)tempRng.Cells[1, 1], i, j);

                                modifiedFormula1 = swapFormulaWithSheetName(Conversions.ToString(worksheet1.get_Range(firstInputRng.get_Address()).Cells[1, 1].offset(i, j).formula), worksheet1.Name);
                                modifiedFormula2 = swapFormulaWithSheetName(Conversions.ToString(worksheet2.get_Range(secondInputRng.get_Address()).Cells[1, 1].offset(i, j).formula), worksheet2.Name);
                                worksheet1.get_Range(firstInputRng.get_Address()).Cells[1, 1].offset(i, j).formula = modifiedFormula2;
                                worksheet2.get_Range(secondInputRng.get_Address()).Cells[1, 1].offset(i, j).formula = modifiedFormula1;

                            }
                        }
                        tempRng.Delete();
                    }

                    else
                    {
                        firstInputRng.ClearFormats();
                        secondInputRng.ClearFormats();

                        for (int i = 0, loopTo4 = firstInputRng.Rows.Count - 1; i <= loopTo4; i++)
                        {
                            for (int j = 0, loopTo5 = firstInputRng.Columns.Count - 1; j <= loopTo5; j++)
                            {

                                modifiedFormula1 = swapFormulaWithSheetName(Conversions.ToString(worksheet1.get_Range(firstInputRng.get_Address()).Cells[1, 1].offset(i, j).formula), worksheet1.Name);
                                modifiedFormula2 = swapFormulaWithSheetName(Conversions.ToString(worksheet2.get_Range(secondInputRng.get_Address()).Cells[1, 1].offset(i, j).formula), worksheet2.Name);
                                worksheet1.get_Range(firstInputRng.get_Address()).Cells[1, 1].offset(i, j).formula = modifiedFormula2;
                                worksheet2.get_Range(secondInputRng.get_Address()).Cells[1, 1].offset(i, j).formula = modifiedFormula1;

                            }
                        }

                    }

                    worksheet1.Activate();
                    firstInputRng.Select();
                }


                else if (radBtnAdjustRef.Checked == true)
                {

                    if (CB_KeepFormatting.Checked == true)
                    {
                        worksheet1.get_Range(firstInputRng.get_Address()).Copy(tempRng);
                        worksheet2.get_Range(secondInputRng.get_Address()).Copy(worksheet1.get_Range(firstInputRng.get_Address()));
                        tempRng.Copy(worksheet2.get_Range(secondInputRng.get_Address()));
                        tempRng.Delete();
                    }

                    else
                    {
                        firstInputRng.ClearFormats();
                        secondInputRng.ClearFormats();

                        worksheet1.get_Range(firstInputRng.get_Address()).Copy(tempRng);
                        worksheet2.get_Range(secondInputRng.get_Address()).Copy(worksheet1.get_Range(firstInputRng.get_Address()));
                        tempRng.Copy(worksheet2.get_Range(secondInputRng.get_Address()));

                        tempRng.Delete();

                    }
                    worksheet1.Activate();
                    firstInputRng.Select();

                }

                Dispose();
            }


            catch (Exception ex)
            {

            }


        }
        public string swapFormulaWithSheetName(string currentFormula, string sheetName)
        {
            string pattern = @"\b([A-Z]+[0-9]+(:[A-Z]+[0-9]+)?)\b";
            string replacement = "";
            char charToFind = ' ';
            int index;

            if (changeState == true)
            {
                if ((worksheet2.Name ?? "") != (worksheet1.Name ?? ""))
                {
                    index = sheetName.IndexOf(charToFind);
                    if (index >= 0)
                    {
                        replacement = "'" + sheetName + "'!$1";
                    }
                    else
                    {
                        replacement = sheetName + "!$1";
                    }
                }

                else
                {
                    replacement = "$1";

                }



            }

            return System.Text.RegularExpressions.Regex.Replace(currentFormula, pattern, replacement);


        }

        public void copyCell(Range destRng, int destOff1, int destOff2, Range srcRng, int srcOff1, int srcOff2)
        {

            destRng.get_Offset(destOff1, destOff2).Font.Name = srcRng.get_Offset(srcOff1, srcOff2).Font.Name;
            destRng.get_Offset(destOff1, destOff2).Font.Size = srcRng.get_Offset(srcOff1, srcOff2).Font.Size;
            destRng.get_Offset(destOff1, destOff2).Font.Color = srcRng.get_Offset(srcOff1, srcOff2).Font.Color;
            destRng.get_Offset(destOff1, destOff2).NumberFormat = srcRng.get_Offset(srcOff1, srcOff2).NumberFormat;
            destRng.get_Offset(destOff1, destOff2).Interior.Color = srcRng.get_Offset(srcOff1, srcOff2).Interior.Color;

            // bold,italic,underline
            destRng.get_Offset(destOff1, destOff2).Font.FontStyle = srcRng.get_Offset(srcOff1, srcOff2).Font.FontStyle;
            destRng.get_Offset(destOff1, destOff2).Font.Underline = srcRng.get_Offset(srcOff1, srcOff2).Font.Underline;




            // border

            destRng.get_Offset(destOff1, destOff2).Borders.LineStyle = srcRng.get_Offset(srcOff1, srcOff2).Borders.LineStyle;
            destRng.get_Offset(destOff1, destOff2).Borders.Weight = srcRng.get_Offset(srcOff1, srcOff2).Borders.Weight;


            // value
            // destRng.Offset(destOff1, destOff2).Value = srcRng.Offset(srcOff1, srcOff2).Value

        }

        private void Form11SwapRanges_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form11SwapRanges_Shown(object sender, EventArgs e)
        {
            Focus();
            BringToFront();
            Activate();
            BeginInvoke(new System.Action(() =>
                {
                    txtSourceRange1.Text = firstInputRng.get_Address();
                    SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                }));
        }

        private void Form11SwapRanges_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }
    }
}