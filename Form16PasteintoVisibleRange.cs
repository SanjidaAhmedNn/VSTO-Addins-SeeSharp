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



    public partial class Form16PasteintoVisibleRange
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
        private Excel.Worksheet outWorksheet;
        private int FocusedTxtBox;
        private Range selectedRange;
        private Range sourceRange, destRange;
        private string WsName, initialWsName;
        private bool changeState = false;
        private bool txtChanged = false;

        public Form16PasteintoVisibleRange()
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

        private void Form16PasteintoVisibleRange_Load(object sender, EventArgs e)
        {

            excelApp = Globals.ThisAddIn.Application;
            workbook = excelApp.ActiveWorkbook;
            worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            // Define a varibale to access a selected range
            Range selectedRng = (Range)excelApp.Selection;

            // Assign the address of selected range that is selcted before loading the form in the textbox "txtSourceRange" 
            // Give foucs to the textbox "txtSourceRange" after the form loads
            txtSourceRange.Text = selectedRng.get_Address();
            txtSourceRange.Focus();


            initialWsName = worksheet.Name;

            KeyPreview = true;

        }

        private void txtSourceRange_TextChanged(object sender, EventArgs e)
        {

            try
            {

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;


                // MsgBox(txtSourceRange1.Text)
                txtChanged = true;

                sourceRange = worksheet.get_Range(txtSourceRange.Text);

                if ((sourceRange.Worksheet.Name ?? "") != (initialWsName ?? ""))
                {

                    txtSourceRange.Text = sourceRange.Worksheet.Name + "!" + sourceRange.get_Address();

                }


                sourceRange.Select();
            }



            catch (Exception ex)
            {

            }

            txtChanged = false;

            txtSourceRange.Focus();


        }
        private void txtDestRange_TextChanged(object sender, EventArgs e)
        {

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                changeState = true;

                txtChanged = true;
                destRange = worksheet.get_Range(txtDestRange.Text);

                destRange.Select();

                if ((destRange.Worksheet.Name ?? "") != (initialWsName ?? ""))
                {

                    txtDestRange.Text = destRange.Worksheet.Name + "!" + destRange.get_Address();

                }
            }


            catch (Exception ex)
            {

            }

            txtChanged = false;
            txtDestRange.Focus();


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
                sourceRange = (Range)excelApp.InputBox("Please Select the First Range", "First Range Selection", selectedRange.get_Address(), Type: 8);
                Show();



                sourceRange.Worksheet.Activate();


                txtSourceRange.Text = sourceRange.get_Address();

                sourceRange.Select();

                txtSourceRange.Focus();
            }



            catch (Exception ex)
            {

                txtSourceRange.Focus();

            }


        }

        private void destinationSelection_Click(object sender, EventArgs e)
        {

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                selectedRange = (Range)excelApp.Selection;
                txtDestRange.Focus();

                Hide();
                destRange = (Range)excelApp.InputBox("Please Select the Second Range", "Second Range Selection", selectedRange.get_Address(), Type: 8);
                Show();


                destRange.Worksheet.Activate();

                // txtDestRange.Text = destRange.Address
                txtDestRange.Text = destRange.get_Address();

                destRange.Select();
                txtDestRange.Focus();
            }




            catch (Exception ex)
            {

                txtDestRange.Focus();

            }




        }



        private void txtSourceRange_GotFocus(object sender, EventArgs e)
        {
            try
            {

                // If txtSourceRange textbox got focus, assign 1 to the global variable "FocusedTxtBox"
                FocusedTxtBox = 1;
            }


            catch (Exception ex)
            {

            }
        }

        private void txtDestRange_GotFocus(object sender, EventArgs e)
        {

            try
            {

                // If txtDestRange textbox got focus, assign 2 to the global variable "FocusedTxtBox"
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

                // checks if the text is changed in ano of the textboxes
                if (txtChanged == false)
                {

                    if (FocusedTxtBox == 1)
                    {
                        txtSourceRange.Text = selectedRange.get_Address();
                        txtSourceRange.Focus();
                    }

                    else if (FocusedTxtBox == 2)
                    {
                        txtDestRange.Text = selectedRange.get_Address();
                    }

                }
            }


            catch (Exception ex)
            {

            }


        }

        private void AutoSelection_Click(object sender, EventArgs e)
        {

            try
            {
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

                sourceRange = selectedRange;
                txtSourceRange.Text = sourceRange.get_Address();
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



            if (string.IsNullOrEmpty(txtSourceRange.Text) & string.IsNullOrEmpty(txtDestRange.Text))
            {

                Interaction.MsgBox("Please select the Source Range and the Destination Range.", MsgBoxStyle.Exclamation, "Error!");
                txtSourceRange.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(txtSourceRange.Text) & !string.IsNullOrEmpty(txtDestRange.Text))
            {

                if (IsValidRng(txtDestRange.Text.ToUpper()) == true)
                {
                    Interaction.MsgBox("Please select data to be copied.", MsgBoxStyle.Exclamation, "Error!");
                    txtSourceRange.Focus();
                    return;
                }
                else
                {
                    Interaction.MsgBox("Please use a valid range in the Destination Range.", MsgBoxStyle.Exclamation, "Error!");
                    txtDestRange.Text = "";
                    txtDestRange.Focus();
                    return;
                }
            }

            else if (string.IsNullOrEmpty(txtDestRange.Text) & !string.IsNullOrEmpty(txtSourceRange.Text))
            {
                if (IsValidRng(txtSourceRange.Text.ToUpper()) == true)
                {
                    Interaction.MsgBox("Please select the Destination Range.", MsgBoxStyle.Exclamation, "Error!");
                    txtDestRange.Focus();
                    return;
                }
                else
                {
                    Interaction.MsgBox("Please select a valid cell range for data to be copied.", MsgBoxStyle.Exclamation, "Error!");
                    txtSourceRange.Text = "";
                    txtSourceRange.Focus();
                    return;
                }
            }

            else if (!string.IsNullOrEmpty(txtSourceRange.Text) & !string.IsNullOrEmpty(txtDestRange.Text))
            {
                if (IsValidRng(txtSourceRange.Text.ToUpper()) == false & IsValidRng(txtDestRange.Text.ToUpper()) == true)
                {
                    Interaction.MsgBox("Please select a valid cell range for data to be copied.", MsgBoxStyle.Exclamation, "Error!");
                    txtSourceRange.Text = "";
                    txtSourceRange.Focus();
                    return;
                }

                else if (IsValidRng(txtSourceRange.Text.ToUpper()) == true & IsValidRng(txtDestRange.Text.ToUpper()) == false)
                {
                    Interaction.MsgBox("Please select a valid cell range for data to be copied.", MsgBoxStyle.Exclamation, "Error!");
                    txtDestRange.Text = "";
                    txtDestRange.Focus();
                    return;
                }
                else if (IsValidRng(txtSourceRange.Text.ToUpper()) == false & IsValidRng(txtDestRange.Text.ToUpper()) == false)
                {
                    Interaction.MsgBox("Please select a valid cell range for data to be copied.", MsgBoxStyle.Exclamation, "Error!");
                    txtSourceRange.Text = "";
                    txtDestRange.Text = "";
                    txtSourceRange.Focus();
                    return;

                }
            }


            int i, j = default, count, pasteValue, lastRowNum, lastColNum;
            int rowNum, colNum;
            int rowOff, colOff;
            string lastRow, lastCol;
            excelApp = Globals.ThisAddIn.Application;
            workbook = excelApp.ActiveWorkbook;
            worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            WsName = worksheet.Name;
            destRange = (Range)destRange.Cells[1, 1];


            if (CB_copyWs.Checked == true)
            {

                workbook.ActiveSheet.Copy(After: workbook.Sheets[workbook.Sheets.Count]);
                outWorksheet = (Excel.Worksheet)workbook.Sheets[workbook.Sheets.Count];


                worksheet = (Excel.Worksheet)workbook.Sheets[WsName];
                worksheet.Activate();


            }


            lastRowNum = 0;
            if (destRange.get_End(XlDirection.xlDown).get_Value() is null)
            {

                while (destRange.get_Offset(lastRowNum, 0).get_Value() is not null)


                    lastRowNum = lastRowNum + 1;

                lastRowNum = destRange.Row + lastRowNum;
            }

            else
            {
                lastRow = destRange.get_End(XlDirection.xlDown).get_Address();
                while (worksheet.get_Range(lastRow).get_Offset(lastRowNum, 0).get_Value() is not null)


                    lastRowNum = lastRowNum + 1;

                lastRowNum = worksheet.get_Range(lastRow).Row + lastRowNum;
            }

            // finding last column number
            lastColNum = 0;
            if (destRange.get_End(XlDirection.xlToRight).get_Value() is null)
            {

                while (destRange.get_Offset(0, lastColNum).get_Value() is not null)


                    lastColNum = lastColNum + 1;

                lastColNum = destRange.Column + lastColNum;
            }

            else
            {
                lastCol = destRange.get_End(XlDirection.xlToRight).get_Address();
                while (worksheet.get_Range(lastCol).get_Offset(0, lastColNum).get_Value() is not null)


                    lastColNum = lastColNum + 1;

                lastColNum = worksheet.get_Range(lastCol).Column + lastColNum;
            }





            // finding the total visible rows
            int visibleRows = 0;
            var loopTo = lastRowNum;
            for (i = destRange.Row; i <= loopTo; i++)
            {

                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.get_Range(worksheet.Cells[i, 1], worksheet.Cells[i, 2]).EntireRow.Hidden, false, false)))
                {
                    visibleRows = visibleRows + 1;
                }


            }
            visibleRows = visibleRows - 1;



            // finding total visible columns
            int visibleCols = 0;
            var loopTo1 = lastColNum;
            for (i = destRange.Column; i <= loopTo1; i++)
            {

                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.get_Range(worksheet.Cells[1, i], worksheet.Cells[2, i]).EntireColumn.Hidden, false, false)))
                {
                    visibleCols = visibleCols + 1;
                }


            }
            visibleCols = visibleCols - 1;


            count = 0;
            pasteValue = 0;

            if (sourceRange.Rows.Count <= visibleRows & sourceRange.Columns.Count <= visibleCols)
            {

                rowOff = 0;

                rowNum = destRange.Row;
                colNum = destRange.Column;

                var loopTo2 = lastRowNum;
                for (i = rowNum; i <= loopTo2; i++)
                {



                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.Cells[i, 1].EntireRow.Hidden, false, false)))
                    {
                        rowOff += 1;
                    }

                    else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.Cells[i, 1].EntireRow.Hidden, true, false)))
                    {

                        goto nextLoop4;
                    }
                    if (rowOff > sourceRange.Rows.Count)
                    {
                        break;
                    }

                    colOff = 0;

                    var loopTo3 = lastColNum;
                    for (j = colNum; j <= loopTo3; j++)
                    {

                        if (colOff + 1 > sourceRange.Columns.Count)
                        {
                            break;
                        }

                        // ' Check if the destination cell (or its row/column) is not hidden.
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.Cells[i, j].EntireColumn.Hidden, false, false)))
                        {
                            colOff += 1;


                            if (CB_keepFormat.Checked == true)
                            {
                                sourceRange.Cells[1, 1].offset((object)(rowOff - 1), (object)(colOff - 1)).Copy(worksheet.Cells[i, j]);
                            }

                            else if (CB_keepFormat.Checked == false)
                            {

                                worksheet.Cells[i, j].value = sourceRange.Cells[1, 1].offset((object)(rowOff - 1), (object)(colOff - 1)).value;

                            }

                        }

                    }

nextLoop4:
                    ;


                }

                worksheet.get_Range(worksheet.Cells[destRange.Row, destRange.Column], worksheet.Cells[i - 1, j - 1]).Select();
            }




            else if (sourceRange.Rows.Count <= visibleRows & sourceRange.Columns.Count > visibleCols)
            {

                rowOff = 0;

                rowNum = destRange.Row;
                colNum = destRange.Column;

                var loopTo6 = lastRowNum;
                for (i = rowNum; i <= loopTo6; i++)
                {



                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.Cells[i, 1].EntireRow.Hidden, false, false)))
                    {
                        rowOff += 1;
                    }

                    else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.Cells[i, 1].EntireRow.Hidden, true, false)))
                    {

                        goto nextLoop;
                    }
                    if (rowOff > sourceRange.Rows.Count)
                    {
                        break;
                    }

                    colOff = 0;

                    var loopTo7 = lastColNum + sourceRange.Columns.Count - visibleCols - 1;
                    for (j = colNum; j <= loopTo7; j++)
                    {



                        if (colOff > sourceRange.Columns.Count)
                        {
                            break;
                        }


                        // ' Check if the destination cell (or its row/column) is not hidden.
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.Cells[i, j].EntireColumn.Hidden, false, false)))
                        {
                            colOff += 1;


                            if (CB_keepFormat.Checked == true)
                            {
                                sourceRange.Cells[1, 1].offset((object)(rowOff - 1), (object)(colOff - 1)).Copy(worksheet.Cells[i, j]);
                            }

                            else if (CB_keepFormat.Checked == false)
                            {

                                worksheet.Cells[i, j].value = sourceRange.Cells[1, 1].offset((object)(rowOff - 1), (object)(colOff - 1)).value;

                            }

                        }

                    }

nextLoop:
                    ;


                }

                worksheet.get_Range(worksheet.Cells[destRange.Row, destRange.Column], worksheet.Cells[i - 1, j - 1]).Select();
            }

            else if (sourceRange.Rows.Count > visibleRows & sourceRange.Columns.Count <= visibleCols)
            {



                rowOff = 0;

                rowNum = destRange.Row;
                colNum = destRange.Column;

                var loopTo8 = lastRowNum + sourceRange.Rows.Count - visibleRows - 1;
                for (i = rowNum; i <= loopTo8; i++)
                {


                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.Cells[i, 1].EntireRow.Hidden, false, false)))
                    {
                        rowOff += 1;
                    }

                    else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.Cells[i, 1].EntireRow.Hidden, true, false)))
                    {

                        goto nextLoop2;
                    }

                    if (rowOff > sourceRange.Rows.Count)
                    {
                        break;
                    }


                    colOff = 0;

                    var loopTo9 = lastColNum;
                    for (j = colNum; j <= loopTo9; j++)
                    {

                        if (colOff > sourceRange.Columns.Count)
                        {
                            break;
                        }


                        // ' Check if the destination cell (or its row/column) is not hidden.
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.Cells[i, j].EntireColumn.Hidden, false, false)))
                        {
                            colOff += 1;


                            if (CB_keepFormat.Checked == true)
                            {
                                sourceRange.Cells[1, 1].offset((object)(rowOff - 1), (object)(colOff - 1)).Copy(worksheet.Cells[i, j]);
                            }

                            else if (CB_keepFormat.Checked == false)
                            {

                                worksheet.Cells[i, j].value = sourceRange.Cells[1, 1].offset((object)(rowOff - 1), (object)(colOff - 1)).value;

                            }

                        }

                    }

nextLoop2:
                    ;


                }

                worksheet.get_Range(worksheet.Cells[destRange.Row, destRange.Column], worksheet.Cells[i - 1, j - 1]).Select();
            }


            else
            {

                rowOff = 0;

                rowNum = destRange.Row;
                colNum = destRange.Column;

                var loopTo4 = lastRowNum + sourceRange.Rows.Count - visibleRows - 1;
                for (i = rowNum; i <= loopTo4; i++)
                {

                    if (rowOff > sourceRange.Rows.Count)
                    {
                        break;
                    }

                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.Cells[i, 1].EntireRow.Hidden, false, false)))
                    {
                        rowOff += 1;
                    }

                    else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.Cells[i, 1].EntireRow.Hidden, true, false)))
                    {

                        goto nextLoop3;
                    }

                    colOff = 0;

                    var loopTo5 = lastColNum + sourceRange.Columns.Count - visibleCols - 1;
                    for (j = colNum; j <= loopTo5; j++)
                    {

                        if (colOff > sourceRange.Columns.Count)
                        {
                            break;
                        }


                        // ' Check if the destination cell (or its row/column) is not hidden.
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.Cells[i, j].EntireColumn.Hidden, false, false)))
                        {
                            colOff += 1;

                            if (CB_keepFormat.Checked == true)
                            {
                                sourceRange.Cells[1, 1].offset((object)(rowOff - 1), (object)(colOff - 1)).Copy(worksheet.Cells[i, j]);
                            }

                            else if (CB_keepFormat.Checked == false)
                            {

                                worksheet.Cells[i, j].value = sourceRange.Cells[1, 1].offset((object)(rowOff - 1), (object)(colOff - 1)).value;

                            }

                        }

                    }

nextLoop3:
                    ;


                }

                worksheet.get_Range(worksheet.Cells[destRange.Row, destRange.Column], worksheet.Cells[i - 1, j - 1]).Select();

            }


            Dispose();





        }
        private void Form16PasteintoVisibleRange_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form16PasteintoVisibleRange_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form16PasteintoVisibleRange_Shown(object sender, EventArgs e)
        {

            Focus();
            BringToFront();
            Activate();
            BeginInvoke(new System.Action(() =>
                {
                    txtSourceRange.Text = sourceRange.get_Address();
                    SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                }));
        }

    }
}

// UNUSED CODES

// While destRange.Offset(count, 0).Value IsNot Nothing

// If destRange.Offset(count, count2).EntireRow.Hidden = False Then
// pasteValue = pasteValue + 1
// count2 = 0
// pasteValue2 = 0

// End If
// If pasteValue > sourceRange.Rows.Count Then
// Exit While
// End If

// While destRange.Offset(count, count2).Value <> Nothing
// If pasteValue2 + 1 > sourceRange.Columns.Count Then
// Exit While
// End If


// If destRange.Offset(count, count2).EntireRow.Hidden = False And destRange.Offset(count, count2).EntireColumn.Hidden = False Then
// pasteValue2 = pasteValue2 + 1


// If CB_keepFormat.Checked = True Then

// 'Call copyCell(destRange, count, count2, worksheet.Range(txtSourceRange.Text).Cells(1, 1), pasteValue - 1, pasteValue2 - 1)
// sourceRange.Cells(1, 1).offset(pasteValue - 1, pasteValue2 - 1).copy(destRange.Cells(1, 1).offset(count, count2))

// Else
// 'sourceRange.Cells(1, 1).offset(pasteValue - 1, pasteValue2 - 1).copy
// 'destRange.Cells(1, 1).offset(count, count2).PasteSpecial(Excel.XlPasteType.xlPasteValues)
// destRange.Offset(count, count2).Value = sourceRange.Cells(1, 1).offset(pasteValue - 1, pasteValue2 - 1).value


// End If


// End If

// count2 += 1

// End While

// count += 1

// End While





// For j = destRange.Row To lastRowNum

// While destRange.Offset(count, 0).Value <> Nothing

// If destRange.Offset(count, count2).EntireRow.Hidden = False Then
// pasteValue = pasteValue + 1
// count2 = 0
// pasteValue2 = 0

// End If
// If pasteValue > sourceRange.Rows.Count Then
// Exit While
// End If

// While destRange.Offset(count, count2).Value <> Nothing
// If pasteValue2 + 1 > sourceRange.Columns.Count Then
// Exit While
// End If


// If destRange.Offset(count, count2).EntireRow.Hidden = False And destRange.Offset(count, count2).EntireColumn.Hidden = False Then
// pasteValue2 = pasteValue2 + 1
// 'If CB_keepFormat.Checked = True Then

// '    Call copyCell(destRange, count, count2, worksheet.Range(txtSourceRange.Text).Cells(1, 1), pasteValue - 1, pasteValue2 - 1)

// '    'Dim borderIndices As Excel.XlBordersIndex() = {Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBordersIndex.xlEdgeTop}
// '    'Dim sourceCell As Excel.Range = sourceRange.Cells(1, 1).Offset(pasteValue - 1, pasteValue2 - 1)
// '    'Dim destCell As Excel.Range = destRange.Cells(1, 1).Offset(count, count2)


// '    'For Each borderIndex As Excel.XlBordersIndex In borderIndices
// '    '    Dim sourceBorder As Excel.Border = sourceCell.Borders(borderIndex)
// '    '    Dim destBorder As Excel.Border = destCell.Borders(borderIndex)

// '    '    If sourceBorder.LineStyle = Excel.XlLineStyle.xlLineStyleNone Then
// '    '        destBorder.LineStyle = Excel.XlLineStyle.xlLineStyleNone
// '    '    Else
// '    '        ' Copying the line style
// '    '        destBorder.LineStyle = sourceBorder.LineStyle

// '    '        ' Copying the color
// '    '        destBorder.Color = sourceBorder.Color

// '    '        ' Copying the weight
// '    '        destBorder.Weight = sourceBorder.Weight

// '    '        ' Copying the TintAndShade
// '    '        destBorder.TintAndShade = sourceBorder.TintAndShade
// '    '    End If

// '    'Next

// 'Else
// '    destRange.Offset(count, count2).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(pasteValue - 1, pasteValue2 - 1).value

// 'End If

// destRange.Offset(count, count2).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(pasteValue - 1, pasteValue2 - 1).value



// End If

// count2 = count2 + 1

// End While

// count = count + 1

// End While

// Next





// Dim count3, count4, count5, l As Integer

// count3 = 0

// For k = lastRowNum To lastRowNum + sourceRange.Rows.Count - visibleRows - 1
// count4 = 0
// count5 = 0
// For l = 1 To lastColNum + sourceRange.Columns.Count - visibleCols - 1

// If worksheet.Cells(lastRowNum, destRange.Column).Offset(count3, l - 1).EntireColumn.Hidden = False Then
// count5 = count5 + 1
// End If


// If count5 > sourceRange.Columns.Count Then
// Exit For
// End If

// If worksheet.Cells(lastRowNum, destRange.Column).Offset(count3, l - 1).EntireColumn.Hidden = False Then

// 'If CB_keepFormat.Checked = True Then

// '    Call copyCell(worksheet.Cells(lastRowNum, destRange.Column), count3, l - 1, worksheet.Range(txtSourceRange.Text).Cells(1, 1), visibleRows + count3, count4)

// '    'Dim borderIndices As Excel.XlBordersIndex() = {Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBordersIndex.xlEdgeTop}
// '    'Dim sourceCell As Excel.Range = sourceRange.Cells(1, 1).Offset(visibleRows + count3, count4)
// '    'Dim destCell As Excel.Range = destRange.Cells(1, 1).Offset(count3, l - 1)


// '    'For Each borderIndex As Excel.XlBordersIndex In borderIndices
// '    '    Dim sourceBorder As Excel.Border = sourceCell.Borders(borderIndex)
// '    '    Dim destBorder As Excel.Border = destCell.Borders(borderIndex)

// '    '    If sourceBorder.LineStyle = Excel.XlLineStyle.xlLineStyleNone Then
// '    '        destBorder.LineStyle = Excel.XlLineStyle.xlLineStyleNone
// '    '    Else
// '    '        ' Copying the line style
// '    '        destBorder.LineStyle = sourceBorder.LineStyle

// '    '        ' Copying the color
// '    '        destBorder.Color = sourceBorder.Color

// '    '        ' Copying the weight
// '    '        destBorder.Weight = sourceBorder.Weight

// '    '        ' Copying the TintAndShade
// '    '        destBorder.TintAndShade = sourceBorder.TintAndShade
// '    '    End If

// '    'Next


// 'Else
// '    worksheet.Cells(lastRowNum, destRange.Column).Offset(count3, l - 1).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(visibleRows + count3, count4).value

// 'End If

// worksheet.Cells(lastRowNum, destRange.Column).Offset(count3, l - 1).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(visibleRows + count3, count4).value



// count4 = count4 + 1
// End If

// Next
// count3 = count3 + 1
// Next




// rowNum = destRange.Row
// colNum = destRange.Column
// count3 = 0
// count4 = visibleCols
// For k = destRange.Row To lastRowNum - 1

// If worksheet.Range(worksheet.Cells(k, 1), worksheet.Cells(k, 2)).EntireRow.Hidden = False Then

// rowNum = worksheet.Range(worksheet.Cells(k, 1), worksheet.Cells(k, 2)).Row

// End If

// If count3 + 1 > sourceRange.Rows.Count Then
// Exit For
// End If


// If Not worksheet.Range(worksheet.Cells(k, 1), worksheet.Cells(k, 2)).EntireRow.Hidden = False And worksheet.Range(worksheet.Cells(k, 1), worksheet.Cells(k + 1, 1)).EntireColumn.Hidden = False Then

// GoTo exitLoop

// End If

// count4 = visibleCols


// For l = lastColNum To lastColNum + sourceRange.Columns.Count - visibleCols - 1


// If worksheet.Range(worksheet.Cells(k, l), worksheet.Cells(k + 1, l)).EntireColumn.Hidden = False Then

// colNum = worksheet.Range(worksheet.Cells(k, l), worksheet.Cells(k + 1, l)).Column

// End If
// If count4 + 1 > sourceRange.Columns.Count Then
// Exit For
// End If


// If worksheet.Range(worksheet.Cells(k, l), worksheet.Cells(k, l + 1)).EntireRow.Hidden = False And worksheet.Range(worksheet.Cells(k, l), worksheet.Cells(k + 1, l)).EntireColumn.Hidden = False Then

// 'If CB_keepFormat.Checked = True Then

// '    Call copyCell(worksheet.Range(worksheet.Cells(rowNum, colNum), worksheet.Cells(rowNum, colNum)), 0, 0, worksheet.Range(txtSourceRange.Text).Cells(1, 1), count3, count4)

// '    'Dim borderIndices As Excel.XlBordersIndex() = {Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBordersIndex.xlEdgeTop}
// '    'Dim sourceCell As Excel.Range = sourceRange.Cells(1, 1).Offset(count3, count4)
// '    'Dim destCell As Excel.Range = destRange.Cells(1, 1).Offset(0, 0)


// '    'For Each borderIndex As Excel.XlBordersIndex In borderIndices
// '    '    Dim sourceBorder As Excel.Border = sourceCell.Borders(borderIndex)
// '    '    Dim destBorder As Excel.Border = destCell.Borders(borderIndex)

// '    '    If sourceBorder.LineStyle = Excel.XlLineStyle.xlLineStyleNone Then
// '    '        destBorder.LineStyle = Excel.XlLineStyle.xlLineStyleNone
// '    '    Else
// '    '        ' Copying the line style
// '    '        destBorder.LineStyle = sourceBorder.LineStyle

// '    '        ' Copying the color
// '    '        destBorder.Color = sourceBorder.Color

// '    '        ' Copying the weight
// '    '        destBorder.Weight = sourceBorder.Weight

// '    '        ' Copying the TintAndShade
// '    '        destBorder.TintAndShade = sourceBorder.TintAndShade
// '    '    End If

// '    'Next


// 'Else
// '    worksheet.Range(worksheet.Cells(rowNum, colNum), worksheet.Cells(rowNum, colNum)).Offset(0, 0).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(count3, count4).value

// 'End If

// worksheet.Range(worksheet.Cells(rowNum, colNum), worksheet.Cells(rowNum, colNum)).Offset(0, 0).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(count3, count4).value



// 'worksheet.Range(worksheet.Cells(rowNum, colNum), worksheet.Cells(rowNum, colNum)).Value = sourceRange.Cells.Offset(count3, count4).Value
// 'sourceRange.Cells.Offset(count3, count4).Copy(worksheet.Cells(rowNum, colNum))

// End If
// count4 = count4 + 1

// Next
// count3 = count3 + 1
// exitLoop:
// Next

// If CB_keepFormat.Checked = True Then
// Dim rowOff, colOff As Integer
// rowOff = 0

// rowNum = destRange.Row
// colNum = destRange.Column

// For i = rowNum To lastRowNum + sourceRange.Rows.Count - visibleRows - 1

// If worksheet.Cells(i, 1).EntireRow.Hidden = False Then
// rowOff += 1

// ElseIf worksheet.Cells(i, 1).EntireRow.Hidden = True Then

// GoTo nextLoop
// End If
// colOff = 0

// For j = colNum To lastColNum + sourceRange.Columns.Count - visibleCols - 1

// 'Dim sourceCell As Excel.Range = sourceRange.Cells(i, j)
// 'Dim destCell As Excel.Range = destRange.Cells(i, j)




// '' Check if the destination cell (or its row/column) is not hidden.
// If worksheet.Cells(i, j).EntireColumn.Hidden = False Then
// colOff += 1
// ' Copy only the formatting.
// sourceRange.Cells(1, 1).offset(rowOff - 1, colOff - 1).Copy()
// worksheet.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)


// End If



// Next
// nextLoop:


// Next


// End If


// worksheet.Range(worksheet.Cells(destRange.Row, destRange.Column), worksheet.Cells(lastRowNum + sourceRange.Rows.Count - visibleRows - 1, lastColNum + sourceRange.Columns.Count - visibleCols - 1)).Select()















// Public Sub copyCell(ByVal destRng As Range, ByVal destOff1 As Integer, ByVal destOff2 As Integer, ByVal srcRng As Range, ByVal srcOff1 As Integer, ByVal srcOff2 As Integer)

// destRng.Offset(destOff1, destOff2).Font.Name = srcRng.Offset(srcOff1, srcOff2).Font.Name
// destRng.Offset(destOff1, destOff2).Font.Size = srcRng.Offset(srcOff1, srcOff2).Font.Size
// destRng.Offset(destOff1, destOff2).Font.Color = srcRng.Offset(srcOff1, srcOff2).Font.Color
// destRng.Offset(destOff1, destOff2).NumberFormat = srcRng.Offset(srcOff1, srcOff2).NumberFormat

// If Not srcRng.Offset(srcOff1, srcOff2).Interior.ColorIndex = -4142 Then

// destRng.Offset(destOff1, destOff2).Interior.Color = srcRng.Offset(srcOff1, srcOff2).Interior.Color

// End If


// 'bold,italic,underline
// destRng.Offset(destOff1, destOff2).Font.FontStyle = srcRng.Offset(srcOff1, srcOff2).Font.FontStyle
// destRng.Offset(destOff1, destOff2).Font.Underline = srcRng.Offset(srcOff1, srcOff2).Font.Underline




// 'border

// Dim borderIndices As Excel.XlBordersIndex() = {Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBordersIndex.xlEdgeTop}
// Dim sourceCell As Excel.Range = sourceRange.Cells(1, 1).Offset(srcOff1, srcOff2)
// Dim destCell As Excel.Range = destRange.Cells(1, 1).Offset(destOff1, destOff2)


// For Each borderIndex As Excel.XlBordersIndex In borderIndices
// Dim sourceBorder As Excel.Border = sourceCell.Borders(borderIndex)
// Dim destBorder As Excel.Border = destCell.Borders(borderIndex)

// If sourceBorder.LineStyle = Excel.XlLineStyle.xlLineStyleNone Then

// destBorder.LineStyle = Excel.XlLineStyle.xlLineStyleNone
// Else
// ' Copying the line style
// destBorder.LineStyle = sourceBorder.LineStyle

// ' Copying the color
// destBorder.Color = sourceBorder.Color

// ' Copying the weight
// destBorder.Weight = sourceBorder.Weight

// ' Copying the TintAndShade
// destBorder.TintAndShade = sourceBorder.TintAndShade
// End If

// Next


// 'value
// destRng.Offset(destOff1, destOff2).Value = srcRng.Offset(srcOff1, srcOff2).Value

// 'gridline
// excelApp.ActiveWindow.DisplayGridlines = True




// End Sub

