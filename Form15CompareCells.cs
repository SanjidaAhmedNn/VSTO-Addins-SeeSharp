using System;
using System.Collections.Generic;
using global::System.ComponentModel;
using global::System.ComponentModel.Design;
using global::System.Drawing;
using System.Linq;
using global::System.Linq.Expressions;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using global::System.Security.Cryptography;
using System.Text;
using global::System.Threading;
using global::System.Windows.Forms;
using static global::System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using global::Microsoft.Office.Interop.Excel;
using Excel = global::Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using global::Microsoft.VisualBasic.Devices;

namespace VSTO_Addins
{

    public partial class Form15CompareCells
    {
        private global::Microsoft.Office.Interop.Excel.Application _excelApp;

        private global::Microsoft.Office.Interop.Excel.Application excelApp
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
        private global::Microsoft.Office.Interop.Excel.Workbook workbook;
        private global::Microsoft.Office.Interop.Excel.Worksheet worksheet, worksheet1, worksheet2;
        private global::Microsoft.Office.Interop.Excel.Worksheet outWorksheet;
        private global::Microsoft.Office.Interop.Excel.Range firstInputRng;
        private global::Microsoft.Office.Interop.Excel.Range secondInputRng;
        private global::System.Int32 FocusedTxtBox;
        private global::Microsoft.Office.Interop.Excel.Range selectedRange;
        private global::System.Int32 firstRngRows, firstRngCols;
        private global::System.Windows.Forms.DialogResult colorPick;
        private global::System.Int32 count;
        private global::System.String rng1CellValue, rng2CellValue, WsName, coloredRng, rngKeyBoard, output, initialWsName;
        private global::System.Boolean changeState = false;
        private global::System.Boolean txtChanged = false;

        public Form15CompareCells()
        {
            InitializeComponent();
        }


        [DllImport("user32")]
        private static extern bool SetWindowPos(global::System.IntPtr hWnd, global::System.IntPtr hWndInsertAfter, global::System.Int32 X, global::System.Int32 Y, global::System.Int32 cx, global::System.Int32 cy, global::System.UInt32 uFlags);
        private const global::System.UInt32 SWP_NOMOVE = 0x2U;
        private const global::System.UInt32 SWP_NOSIZE = 0x1U;
        private const global::System.UInt32 SWP_NOACTIVATE = 0x10U;
        private const global::System.Int32 HWND_TOPMOST = -(1);

        private void Form1_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
            {
                this.btnOK.PerformClick();
            }
        }

        private void txtSourceRange1_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {

                this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
                this.workbook = excelApp.ActiveWorkbook;
                this.worksheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;


                // MsgBox(txtSourceRange1.Text)
                this.txtChanged = true;
                this.firstInputRng = this.worksheet.get_Range(this.txtSourceRange1.Text);


                this.lblSourceRng1.Text = (((("1st Source Range (" + this.firstInputRng.Rows.Count) + " rows x ") + this.firstInputRng.Columns.Count) + " columns)");

                this.firstInputRng.Select();

                this.firstRngRows = this.worksheet.get_Range(this.txtSourceRange1.Text).Rows.Count;
                this.firstRngCols = this.worksheet.get_Range(this.txtSourceRange1.Text).Columns.Count;

                if (((firstInputRng.Worksheet.Name ?? "") != (this.initialWsName ?? "")))
                {

                    this.txtSourceRange1.Text = ((firstInputRng.Worksheet.Name + "!") + this.firstInputRng.get_Address());

                }
            }

            // If secondInputRng.Worksheet.Name <> firstInputRng.Worksheet.Name Then

            // txtSourceRange2.Text = secondInputRng.Worksheet.Name & "!" & secondInputRng.Address
            // secondInputRng = worksheet.Range(Microsoft.VisualBasic.Right(txtSourceRange2.Text, Len(txtSourceRange2.Text) - txtSourceRange2.Text.IndexOf("!") - 1))
            // lblSourceRng2.Text = "2nd Source Range (" & secondInputRng.Rows.Count & " rows x " & secondInputRng.Columns.Count & " columns)"
            // Else
            // txtSourceRange2.Text = secondInputRng.Address
            // lblSourceRng2.Text = "2nd Source Range (" & secondInputRng.Rows.Count & " rows x " & secondInputRng.Columns.Count & " columns)"
            // End If

            catch (global::System.Exception ex)
            {

            }

            this.Display();
            this.txtChanged = false;

            this.txtSourceRange1.Focus();

        }



        private void txtSourceRange2_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
                this.workbook = excelApp.ActiveWorkbook;
                this.worksheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;

                this.changeState = true;

                this.txtChanged = true;
                this.secondInputRng = this.worksheet.get_Range(this.txtSourceRange2.Text);

                this.lblSourceRng2.Text = (((("2nd Source Range (" + this.secondInputRng.Rows.Count) + " rows x ") + this.secondInputRng.Columns.Count) + " columns)");

                this.secondInputRng.Select();

                if (((secondInputRng.Worksheet.Name ?? "") != (this.initialWsName ?? "")))
                {

                    this.txtSourceRange2.Text = ((secondInputRng.Worksheet.Name + "!") + this.secondInputRng.get_Address());


                }
            }


            // If secondInputRng.Worksheet.Name <> firstInputRng.Worksheet.Name Then

            // txtSourceRange2.Text = secondInputRng.Worksheet.Name & "!" & secondInputRng.Address
            // '    secondInputRng = worksheet.Range(Microsoft.VisualBasic.Right(txtSourceRange2.Text, Len(txtSourceRange2.Text) - txtSourceRange2.Text.IndexOf("!") - 1))
            // '    lblSourceRng2.Text = "2nd Source Range (" & secondInputRng.Rows.Count & " rows x " & secondInputRng.Columns.Count & " columns)"
            // 'Else
            // '    txtSourceRange2.Text = secondInputRng.Address
            // '    lblSourceRng2.Text = "2nd Source Range (" & secondInputRng.Rows.Count & " rows x " & secondInputRng.Columns.Count & " columns)"

            // End If

            catch (global::System.Exception ex)
            {

            }

            this.Display();
            this.txtChanged = false;
            this.txtSourceRange2.Focus();

        }

        private void Form15CompareCells_Load(global::System.Object sender, global::System.EventArgs e)
        {

            this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
            this.workbook = excelApp.ActiveWorkbook;
            this.worksheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;

            global::Microsoft.Office.Interop.Excel.Range selectedRng = (global::Microsoft.Office.Interop.Excel.Range)excelApp.Selection;

            this.txtSourceRange1.Focus();
            this.txtSourceRange1.Text = selectedRng.get_Address();

            this.radBtnSameValues.Checked = true;

            this.initialWsName = worksheet.Name;

            this.KeyPreview = true;

        }

        private void rngSelection1_Click(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {

                this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
                this.workbook = excelApp.ActiveWorkbook;
                this.worksheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                this.selectedRange = (global::Microsoft.Office.Interop.Excel.Range)excelApp.Selection;
                this.txtSourceRange1.Focus();

                this.Hide();
                this.firstInputRng = (global::Microsoft.Office.Interop.Excel.Range)this.excelApp.InputBox("Please Select the First Range", "First Range Selection", this.selectedRange.get_Address(), Type: (global::System.Object)8);
                this.Show();

                this.firstInputRng.Worksheet.Activate();

                this.txtSourceRange1.Text = this.firstInputRng.get_Address();

                this.firstInputRng.Select();

                this.txtSourceRange1.Focus();
            }

            catch (global::System.Exception ex)
            {

                this.txtSourceRange1.Focus();

            }

        }

        private void rngSelection2_Click(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
                this.workbook = excelApp.ActiveWorkbook;
                this.worksheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                this.selectedRange = (global::Microsoft.Office.Interop.Excel.Range)excelApp.Selection;
                this.txtSourceRange2.Focus();

                this.Hide();
                this.secondInputRng = (global::Microsoft.Office.Interop.Excel.Range)this.excelApp.InputBox("Please Select the Second Range", "Second Range Selection", this.selectedRange.get_Address(), Type: (global::System.Object)8);
                this.Show();

                this.secondInputRng.Worksheet.Activate();


                this.txtSourceRange2.Text = this.secondInputRng.get_Address();

                this.secondInputRng.Select();
                this.txtSourceRange2.Focus();
            }

            catch (global::System.Exception ex)
            {

                this.txtSourceRange2.Focus();

            }
        }

        private void AutoSelection1_Click(global::System.Object sender, global::System.EventArgs e)
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

                // Call Display()

                // txtSourceRange1.Text = selectedRange.Worksheet.Name & "!" & selectedRange.Address


                this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
                this.workbook = excelApp.ActiveWorkbook;
                this.worksheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                this.selectedRange = (global::Microsoft.Office.Interop.Excel.Range)excelApp.Selection;

                var activeRange = excelApp.ActiveCell;

                global::System.Int32 startRow = activeRange.Row;
                global::System.Int32 startColumn = activeRange.Column;
                global::System.Int32 endRow = activeRange.Row;
                global::System.Int32 endColumn = activeRange.Column;

                // Find the upper boundary
                while ((((startRow) > (1)) && (!((worksheet.Cells[(global::System.Object)((startRow) - (1)), (global::System.Object)startColumn].Value == null)))))
                    startRow -= 1;

                // Find the lower boundary
                while (!((worksheet.Cells[(global::System.Object)((endRow) + (1)), (global::System.Object)endColumn].Value == null)))
                    endRow += 1;

                // Find the left boundary
                while ((((startColumn) > (1)) && (!((worksheet.Cells[(global::System.Object)startRow, (global::System.Object)((startColumn) - (1))].Value == null)))))
                    startColumn -= 1;

                // Find the right boundary
                while (!((worksheet.Cells[(global::System.Object)endRow, (global::System.Object)((endColumn) + (1))].Value == null)))
                    endColumn += 1;

                // Select the determined range
                this.worksheet.get_Range(worksheet.Cells[(global::System.Object)startRow, (global::System.Object)startColumn], worksheet.Cells[(global::System.Object)endRow, (global::System.Object)endColumn]).Select();

                this.firstInputRng = this.selectedRange;
                this.txtSourceRange1.Text = this.firstInputRng.get_Address();

                this.firstRngRows = this.selectedRange.Rows.Count;
                this.firstRngCols = this.selectedRange.Columns.Count;
            }



            catch (global::System.Exception ex)
            {

            }

        }

        private void AutoSelection2_Click(global::System.Object sender, global::System.EventArgs e)
        {

            global::Microsoft.Office.Interop.Excel.Range firstCell;

            this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
            this.workbook = excelApp.ActiveWorkbook;
            this.worksheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
            this.selectedRange = (global::Microsoft.Office.Interop.Excel.Range)excelApp.Selection;
            this.selectedRange.Select();

            global::System.String bottomRight;
            firstCell = (global::Microsoft.Office.Interop.Excel.Range)this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1];

            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1].Offset((global::System.Object)1, (global::System.Object)0).Value, null, false)))
            {

                for (global::System.Int32 i = 0, loopTo = (this.firstRngCols) - (1); i <= loopTo; i++)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1].offset((global::System.Object)0, i).value, null, false)))
                    {
                        this.selectedRange = this.worksheet.get_Range(this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1], this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1].Offset((global::System.Object)0, i));
                    }
                    this.selectedRange.Select();
                }
            }

            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1].Offset((global::System.Object)0, (global::System.Object)1).Value, null, false)))
            {
                for (global::System.Int32 i = 0, loopTo1 = (this.firstRngRows) - (1); i <= loopTo1; i++)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1].offset(i, (global::System.Object)0).value, null, false)))
                    {
                        this.selectedRange = this.worksheet.get_Range(this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1], this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1].Offset(i, (global::System.Object)0));
                    }
                    this.selectedRange.Select();
                }
            }

            else
            {

                bottomRight = firstCell.get_End(global::Microsoft.Office.Interop.Excel.XlDirection.xlToRight).get_Address();
                bottomRight = this.worksheet.get_Range(bottomRight).get_End(global::Microsoft.Office.Interop.Excel.XlDirection.xlDown).get_Address();

                this.selectedRange = this.worksheet.get_Range(firstCell, this.worksheet.get_Range(bottomRight));

                if ((((this.selectedRange.Rows.Count) == (1)) & ((this.selectedRange.Columns.Count) >= (this.firstRngCols))))
                {
                    this.selectedRange = this.worksheet.get_Range(this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1], this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1].Offset((global::System.Object)0, (global::System.Object)((this.firstRngCols) - (1))));
                    this.selectedRange.Select();
                }

                else if ((((this.selectedRange.Rows.Count) == (1)) & ((this.selectedRange.Columns.Count) < (this.firstRngCols))))
                {
                    this.selectedRange = this.worksheet.get_Range(this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1], this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1].Offset((global::System.Object)0, (global::System.Object)((this.selectedRange.Columns.Count) - (1))));
                    this.selectedRange.Select();
                }

                else if ((((this.selectedRange.Columns.Count) == (1)) & ((this.selectedRange.Rows.Count) >= (this.firstRngRows))))
                {
                    this.selectedRange = this.worksheet.get_Range(this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1], this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1].Offset((global::System.Object)((this.firstRngRows) - (1)), (global::System.Object)0));
                    this.selectedRange.Select();
                }

                else if ((((this.selectedRange.Columns.Count) == (1)) & ((this.selectedRange.Rows.Count) < (this.firstRngRows))))
                {
                    this.selectedRange = this.worksheet.get_Range(this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1], this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1].Offset((global::System.Object)((this.selectedRange.Rows.Count) - (1)), (global::System.Object)0));
                    this.selectedRange.Select();
                }


                else
                {
                    bottomRight = firstCell.get_End(global::Microsoft.Office.Interop.Excel.XlDirection.xlToRight).get_Address();
                    bottomRight = this.worksheet.get_Range(bottomRight).get_End(global::Microsoft.Office.Interop.Excel.XlDirection.xlDown).get_Address();

                    this.selectedRange = this.worksheet.get_Range(firstCell, this.worksheet.get_Range(bottomRight));

                    if ((((this.selectedRange.Rows.Count) == (this.firstRngRows)) & ((this.selectedRange.Columns.Count) == (this.firstRngCols))))
                    {
                        firstCell = (global::Microsoft.Office.Interop.Excel.Range)this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1];
                        this.selectedRange = this.worksheet.get_Range(firstCell.get_Offset((global::System.Object)0, (global::System.Object)0), firstCell.get_Offset((global::System.Object)((this.firstRngRows) - (1)), (global::System.Object)((this.firstRngCols) - (1))));
                        this.selectedRange.Select();
                    }

                    else if ((((this.selectedRange.Rows.Count) == (this.firstRngRows)) & ((this.selectedRange.Columns.Count) > (this.firstRngCols))))
                    {
                        firstCell = (global::Microsoft.Office.Interop.Excel.Range)this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1];
                        this.selectedRange = this.worksheet.get_Range(firstCell.get_Offset((global::System.Object)0, (global::System.Object)0), firstCell.get_Offset((global::System.Object)((this.firstRngRows) - (1)), (global::System.Object)((this.firstRngCols) - (1))));
                        this.selectedRange.Select();
                    }

                    else if ((((this.selectedRange.Rows.Count) == (this.firstRngRows)) & ((this.selectedRange.Columns.Count) < (this.firstRngCols))))
                    {
                        firstCell = (global::Microsoft.Office.Interop.Excel.Range)this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1];
                        bottomRight = firstCell.get_End(global::Microsoft.Office.Interop.Excel.XlDirection.xlToRight).get_Address();
                        bottomRight = this.worksheet.get_Range(bottomRight).get_End(global::Microsoft.Office.Interop.Excel.XlDirection.xlDown).get_Address();

                        this.selectedRange = this.worksheet.get_Range(firstCell.get_Offset((global::System.Object)0, (global::System.Object)0), this.worksheet.get_Range(bottomRight));
                        this.selectedRange.Select();
                    }

                    else if ((((this.selectedRange.Rows.Count) > (this.firstRngRows)) & ((this.selectedRange.Columns.Count) == (this.firstRngCols))))
                    {
                        firstCell = (global::Microsoft.Office.Interop.Excel.Range)this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1];
                        this.selectedRange = this.worksheet.get_Range(firstCell.get_Offset((global::System.Object)0, (global::System.Object)0), firstCell.get_Offset((global::System.Object)((this.firstRngRows) - (1)), (global::System.Object)((this.firstRngCols) - (1))));
                        this.selectedRange.Select();
                    }

                    else if ((((this.selectedRange.Rows.Count) > (this.firstRngRows)) & ((this.selectedRange.Columns.Count) > (this.firstRngCols))))
                    {
                        firstCell = (global::Microsoft.Office.Interop.Excel.Range)this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1];
                        this.selectedRange = this.worksheet.get_Range(firstCell.get_Offset((global::System.Object)0, (global::System.Object)0), firstCell.get_Offset((global::System.Object)((this.firstRngRows) - (1)), (global::System.Object)((this.firstRngCols) - (1))));
                        this.selectedRange.Select();
                    }

                    else if ((((this.selectedRange.Rows.Count) > (this.firstRngRows)) & ((this.selectedRange.Columns.Count) < (this.firstRngCols))))
                    {
                        firstCell = (global::Microsoft.Office.Interop.Excel.Range)this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1];
                        bottomRight = firstCell.get_End(global::Microsoft.Office.Interop.Excel.XlDirection.xlToRight).get_Address();
                        bottomRight = this.worksheet.get_Range(bottomRight).get_Offset((global::System.Object)((this.firstRngRows) - (1)), (global::System.Object)0).get_Address();

                        this.selectedRange = this.worksheet.get_Range(firstCell.get_Offset((global::System.Object)0, (global::System.Object)0), this.worksheet.get_Range(bottomRight));
                        this.selectedRange.Select();
                    }

                    else if ((((this.selectedRange.Rows.Count) < (this.firstRngRows)) & ((this.selectedRange.Columns.Count) == (this.firstRngCols))))
                    {
                        firstCell = (global::Microsoft.Office.Interop.Excel.Range)this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1];
                        bottomRight = firstCell.get_End(global::Microsoft.Office.Interop.Excel.XlDirection.xlToRight).get_Address();
                        bottomRight = this.worksheet.get_Range(bottomRight).get_End(global::Microsoft.Office.Interop.Excel.XlDirection.xlDown).get_Address();

                        this.selectedRange = this.worksheet.get_Range(firstCell.get_Offset((global::System.Object)0, (global::System.Object)0), this.worksheet.get_Range(bottomRight));
                        this.selectedRange.Select();
                    }
                    else if ((((this.selectedRange.Rows.Count) < (this.firstRngRows)) & ((this.selectedRange.Columns.Count) > (this.firstRngCols))))
                    {

                        firstCell = (global::Microsoft.Office.Interop.Excel.Range)this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1];
                        bottomRight = firstCell.get_Offset((global::System.Object)0, (global::System.Object)((this.firstRngCols) - (1))).get_Address();
                        bottomRight = this.worksheet.get_Range(bottomRight).get_End(global::Microsoft.Office.Interop.Excel.XlDirection.xlDown).get_Address();

                        this.selectedRange = this.worksheet.get_Range(firstCell.get_Offset((global::System.Object)0, (global::System.Object)0), this.worksheet.get_Range(bottomRight));
                        this.selectedRange.Select();
                    }


                    else if ((((this.selectedRange.Rows.Count) < (this.firstRngRows)) & ((this.selectedRange.Columns.Count) < (this.firstRngCols))))
                    {
                        firstCell = (global::Microsoft.Office.Interop.Excel.Range)this.selectedRange.Cells[(global::System.Object)1, (global::System.Object)1];
                        bottomRight = firstCell.get_End(global::Microsoft.Office.Interop.Excel.XlDirection.xlToRight).get_Address();
                        bottomRight = this.worksheet.get_Range(bottomRight).get_End(global::Microsoft.Office.Interop.Excel.XlDirection.xlDown).get_Address();

                        this.selectedRange = this.worksheet.get_Range(firstCell.get_Offset((global::System.Object)0, (global::System.Object)0), this.worksheet.get_Range(bottomRight));
                        this.selectedRange.Select();

                    }
                }

            }

            this.secondInputRng = this.selectedRange;
            this.txtSourceRange2.Text = this.secondInputRng.get_Address();


        }

        private void txtSourceRange1_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {

                this.FocusedTxtBox = 1;
            }
            // Call Display()

            catch (global::System.Exception ex)
            {

            }
        }
        private void txtSourceRange2_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {

                this.FocusedTxtBox = 2;
            }
            // Call Display()

            catch (global::System.Exception ex)
            {

            }
        }

        private void Form1_Activated(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {

                this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;

                this.excelApp.SheetSelectionChange += this.rngSelectionFromTxtBox;
            }

            catch (global::System.Exception ex)
            {

            }

        }
        private void rngSelectionFromTxtBox(global::System.Object Sh, global::Microsoft.Office.Interop.Excel.Range Target)
        {

            try
            {

                this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
                this.worksheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                this.selectedRange = (global::Microsoft.Office.Interop.Excel.Range)excelApp.Selection;
                this.selectedRange.Select();


                if (((this.txtChanged) == (false)))
                {


                    if (((this.FocusedTxtBox) == (1)))
                    {
                        this.txtSourceRange1.Text = this.selectedRange.get_Address();
                        this.txtSourceRange1.Focus();
                    }

                    else if (((this.FocusedTxtBox) == (2)))
                    {
                        this.txtSourceRange2.Text = this.selectedRange.get_Address();
                    }

                }
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void btnCanecl_Click(global::System.Object sender, global::System.EventArgs e)
        {
            this.Dispose();
        }
        public global::System.Boolean IsValidRng(global::System.String input)
        {
            // "^(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$"

            global::System.String pattern = @"^(.*!)?(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$";
            return global::System.Text.RegularExpressions.Regex.IsMatch(input, pattern);

        }

        private void btnOK_Click(global::System.Object sender, global::System.EventArgs e)
        {


            if ((string.IsNullOrEmpty(this.txtSourceRange1.Text) & string.IsNullOrEmpty(this.txtSourceRange2.Text)))
            {

                global::Microsoft.VisualBasic.Interaction.MsgBox("Please select the first and the second range.", global::Microsoft.VisualBasic.MsgBoxStyle.Exclamation, "Error!");
                this.txtSourceRange1.Focus();
                return;
            }
            else if ((string.IsNullOrEmpty(this.txtSourceRange1.Text) & !string.IsNullOrEmpty(this.txtSourceRange2.Text)))
            {

                if (((this.IsValidRng(this.txtSourceRange2.Text.ToUpper())) == (true)))
                {
                    global::Microsoft.VisualBasic.Interaction.MsgBox("Please select the first range.", global::Microsoft.VisualBasic.MsgBoxStyle.Exclamation, "Error!");
                    this.txtSourceRange1.Focus();
                    return;
                }
                else
                {
                    global::Microsoft.VisualBasic.Interaction.MsgBox("Please use a valid range in the 2nd Source Range.", global::Microsoft.VisualBasic.MsgBoxStyle.Exclamation, "Error!");
                    this.txtSourceRange2.Text = "";
                    this.txtSourceRange2.Focus();
                    return;
                }
            }

            else if ((string.IsNullOrEmpty(this.txtSourceRange2.Text) & !string.IsNullOrEmpty(this.txtSourceRange1.Text)))
            {
                if (((this.IsValidRng(this.txtSourceRange1.Text.ToUpper())) == (true)))
                {
                    global::Microsoft.VisualBasic.Interaction.MsgBox("Please select the second range.", global::Microsoft.VisualBasic.MsgBoxStyle.Exclamation, "Error!");
                    this.txtSourceRange2.Focus();
                    return;
                }
                else
                {
                    global::Microsoft.VisualBasic.Interaction.MsgBox("Please use a valid range in the 1st Source Range.", global::Microsoft.VisualBasic.MsgBoxStyle.Exclamation, "Error!");
                    this.txtSourceRange1.Text = "";
                    this.txtSourceRange1.Focus();
                    return;
                }
            }

            else if ((!string.IsNullOrEmpty(this.txtSourceRange1.Text) & !string.IsNullOrEmpty(this.txtSourceRange2.Text)))
            {
                if ((((this.IsValidRng(this.txtSourceRange1.Text.ToUpper())) == (false)) & ((this.IsValidRng(this.txtSourceRange2.Text.ToUpper())) == (true))))
                {
                    global::Microsoft.VisualBasic.Interaction.MsgBox("Please use a valid range in the 1st Source Range.", global::Microsoft.VisualBasic.MsgBoxStyle.Exclamation, "Error!");
                    this.txtSourceRange1.Text = "";
                    this.txtSourceRange1.Focus();
                    return;
                }

                else if ((((this.IsValidRng(this.txtSourceRange1.Text.ToUpper())) == (true)) & ((this.IsValidRng(this.txtSourceRange2.Text.ToUpper())) == (false))))
                {
                    global::Microsoft.VisualBasic.Interaction.MsgBox("Please use a valid range in the 2nd Source Range.", global::Microsoft.VisualBasic.MsgBoxStyle.Exclamation, "Error!");
                    this.txtSourceRange2.Text = "";
                    this.txtSourceRange2.Focus();
                    return;
                }
                else if ((((this.IsValidRng(this.txtSourceRange1.Text.ToUpper())) == (false)) & ((this.IsValidRng(this.txtSourceRange2.Text.ToUpper())) == (false))))
                {
                    global::Microsoft.VisualBasic.Interaction.MsgBox("Please use a valid range in the Source Ranges.", global::Microsoft.VisualBasic.MsgBoxStyle.Exclamation, "Error!");
                    this.txtSourceRange1.Text = "";
                    this.txtSourceRange2.Text = "";
                    this.txtSourceRange1.Focus();
                    return;

                }
            }

            if ((((this.firstInputRng.Rows.Count) != (this.secondInputRng.Rows.Count)) & ((this.firstInputRng.Columns.Count) != (this.secondInputRng.Columns.Count))))
            {

                global::Microsoft.VisualBasic.Interaction.MsgBox("You must use same number of rows and columns in both ranges.", Title: "Warning!");
                this.txtSourceRange2.Focus();
                return;
            }

            else if ((((this.firstInputRng.Rows.Count) != (this.secondInputRng.Rows.Count)) & ((this.firstInputRng.Columns.Count) == (this.secondInputRng.Columns.Count))))
            {
                global::Microsoft.VisualBasic.Interaction.MsgBox("Please match the source range row size.", Title: "Warning!");
                this.txtSourceRange2.Focus();
                // Me.Dispose()
                return;
            }
            else if ((((this.firstInputRng.Rows.Count) == (this.secondInputRng.Rows.Count)) & ((this.firstInputRng.Columns.Count) != (this.secondInputRng.Columns.Count))))
            {
                global::Microsoft.VisualBasic.Interaction.MsgBox("Please match the source range column size.", Title: "Warning!");
                this.txtSourceRange2.Focus();
                return;

            }

            this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
            global::System.Int32 i, j;
            global::System.String rng1CellValue, rng2CellValue;
            global::System.String coloredRng;
            global::System.String temp;

            this.worksheet1 = this.firstInputRng.Worksheet;
            this.worksheet2 = this.secondInputRng.Worksheet;


            this.count = 0;
            coloredRng = "";
            temp = this.txtSourceRange2.Text;


            if (((this.checkBoxCopyWs.Checked) == (true)))
            {

                this.worksheet1.Copy(After: workbook.Sheets[(global::System.Object)workbook.Sheets.Count]);
                this.outWorksheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[(global::System.Object)workbook.Sheets.Count];

                this.worksheet2.Activate();

                this.txtSourceRange2.Text = temp;

            }


            if (((this.checkBoxFormatting.Checked) == (false)))
            {

                this.firstInputRng.ClearFormats();

            }



            if (((this.radBtnSameValues.Checked) == (true)))
            {
                if (((this.checkBoxCase.Checked) == (true)))
                {

                    // 1st Range >> 2nd Range >> radbtnSameValues checked >> case sensitive checked >> fill/font color both are selected >> OK
                    if ((((this.checkBoxFillBack.Checked) == (true)) & ((this.checkBoxFillFont.Checked) == (true))))
                    {
                        var loopTo = this.firstInputRng.Rows.Count;
                        for (i = 1; i <= loopTo; i++)
                        {
                            var loopTo1 = this.firstInputRng.Columns.Count;
                            for (j = 1; j <= loopTo1; j++)
                            {

                                rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                                if ((rng1CellValue is null & rng2CellValue is null))
                                {
                                    this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)this.CBFillBackground.BackColor;

                                    this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)this.CbFillFont.BackColor;
                                    this.count = ((this.count) + (1));
                                    coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));

                                    goto nextLoop1;
                                }

                                else if ((rng1CellValue is null | rng2CellValue is null))
                                {
                                    goto nextLoop1;

                                }

                                // handles comparison of different type o variables
                                if ((global::Microsoft.VisualBasic.Information.VarType(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value) != global::Microsoft.VisualBasic.Information.VarType(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                {
                                    goto nextLoop1;
                                }

                                else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value, this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value, false)))
                                {

                                    this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)this.CBFillBackground.BackColor;

                                    this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)this.CbFillFont.BackColor;
                                    this.count = ((this.count) + (1));
                                    coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));
                                }

nextLoop1:
                                ;

                            }
                        }
                    }

                    // 1st Range >> 2nd Range >> radbtnSameValues checked >> case sensitive checked >> only fill color is selected >> OK
                    else if ((((this.checkBoxFillBack.Checked) == (true)) & ((this.checkBoxFillFont.Checked) == (false))))
                    {


                        var loopTo4 = this.firstInputRng.Rows.Count;
                        for (i = 1; i <= loopTo4; i++)
                        {
                            var loopTo5 = this.firstInputRng.Columns.Count;
                            for (j = 1; j <= loopTo5; j++)
                            {

                                rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                                if ((rng1CellValue is null & rng2CellValue is null))
                                {
                                    this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)this.CBFillBackground.BackColor;
                                    this.count = ((this.count) + (1));
                                    coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));

                                    goto nextLoop2;
                                }

                                else if ((rng1CellValue is null | rng2CellValue is null))
                                {
                                    goto nextLoop2;

                                }

                                if ((global::Microsoft.VisualBasic.Information.VarType(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value) != global::Microsoft.VisualBasic.Information.VarType(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                {
                                    goto nextLoop2;
                                }

                                else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value, this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value, false)))
                                {

                                    this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)this.CBFillBackground.BackColor;
                                    this.count = ((this.count) + (1));
                                    coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));
                                }

nextLoop2:
                                ;

                            }
                        }
                    }

                    // 1st Range >> 2nd Range >> radbtnSameValues checked >> case sensitive checked >> only font color is selected >> OK
                    else if ((((this.checkBoxFillBack.Checked) == (false)) & ((this.checkBoxFillFont.Checked) == (true))))
                    {

                        var loopTo6 = this.firstInputRng.Rows.Count;
                        for (i = 1; i <= loopTo6; i++)
                        {
                            var loopTo7 = this.firstInputRng.Columns.Count;
                            for (j = 1; j <= loopTo7; j++)
                            {


                                rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                                if ((rng1CellValue is null & rng2CellValue is null))
                                {
                                    this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)this.CbFillFont.BackColor;

                                    this.count = ((this.count) + (1));
                                    coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));

                                    goto nextLoop3;
                                }

                                else if ((rng1CellValue is null | rng2CellValue is null))
                                {
                                    goto nextLoop3;

                                }

                                if ((global::Microsoft.VisualBasic.Information.VarType(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value) != global::Microsoft.VisualBasic.Information.VarType(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                {
                                    goto nextLoop3;
                                }

                                else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value, this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value, false)))
                                {

                                    this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)this.CbFillFont.BackColor;

                                    this.count = ((this.count) + (1));
                                    coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));
                                }

nextLoop3:
                                ;

                            }
                        }
                    }

                    // 1st Range >> 2nd Range >> radbtnSameValues checked >> case sensitive checked >> fill/font color is not selected >> OK
                    else
                    {

                        var loopTo2 = this.firstInputRng.Rows.Count;
                        for (i = 1; i <= loopTo2; i++)
                        {
                            var loopTo3 = this.firstInputRng.Columns.Count;
                            for (j = 1; j <= loopTo3; j++)
                            {

                                rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                                if ((rng1CellValue is null & rng2CellValue is null))
                                {
                                    this.count = ((this.count) + (1));
                                    coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));

                                    goto nextLoop4;
                                }

                                else if ((rng1CellValue is null | rng2CellValue is null))
                                {
                                    goto nextLoop4;

                                }


                                // If variable type of two compared cell are different
                                if ((global::Microsoft.VisualBasic.Information.VarType(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value) != global::Microsoft.VisualBasic.Information.VarType(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                {
                                    goto nextLoop4;
                                }

                                else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value, this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value, false)))
                                {
                                    this.count = ((this.count) + (1));
                                    coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));
                                }

nextLoop4:
                                ;

                            }
                        }

                    }
                }

                // 1st Range >> 2nd Range >> radbtnSameValues checked >> case sensitive unchecked >> fill/font color both are selected >> OK
                else if ((((this.checkBoxFillBack.Checked) == (true)) & ((this.checkBoxFillFont.Checked) == (true))))
                {
                    var loopTo8 = this.firstInputRng.Rows.Count;
                    for (i = 1; i <= loopTo8; i++)
                    {
                        var loopTo9 = this.firstInputRng.Columns.Count;
                        for (j = 1; j <= loopTo9; j++)
                        {
                            rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                            rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                            if ((rng1CellValue is null & rng2CellValue is null))
                            {
                                this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)this.CBFillBackground.BackColor;

                                this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)this.CbFillFont.BackColor;
                                this.count = ((this.count) + (1));
                                coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));

                                goto nextLoop5;
                            }

                            else if ((rng1CellValue is null | rng2CellValue is null))
                            {
                                goto nextLoop5;

                            }

                            if ((global::Microsoft.VisualBasic.Information.VarType(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value) != global::Microsoft.VisualBasic.Information.VarType(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                            {
                                goto nextLoop5;
                            }

                            else if (((rng1CellValue.ToUpper() ?? "") == (rng2CellValue.ToUpper() ?? "")))
                            {

                                this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)this.CBFillBackground.BackColor;

                                this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)this.CbFillFont.BackColor;
                                this.count = ((this.count) + (1));
                                coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));
                            }

nextLoop5:
                            ;

                        }
                    }
                }

                // 1st Range >> 2nd Range >> radbtnSameValues checked >> case sensitive unchecked >> only fill color is selected >> OK
                else if ((((this.checkBoxFillBack.Checked) == (true)) & ((this.checkBoxFillFont.Checked) == (false))))
                {
                    var loopTo12 = this.firstInputRng.Rows.Count;
                    for (i = 1; i <= loopTo12; i++)
                    {
                        var loopTo13 = this.firstInputRng.Columns.Count;
                        for (j = 1; j <= loopTo13; j++)
                        {
                            rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                            rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                            rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                            rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                            if ((rng1CellValue is null & rng2CellValue is null))
                            {
                                this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)this.CBFillBackground.BackColor;

                                this.count = ((this.count) + (1));
                                coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));

                                goto nextLoop6;
                            }

                            else if ((rng1CellValue is null | rng2CellValue is null))
                            {
                                goto nextLoop6;

                            }


                            if ((global::Microsoft.VisualBasic.Information.VarType(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value) != global::Microsoft.VisualBasic.Information.VarType(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                            {
                                goto nextLoop6;
                            }

                            else if (((rng1CellValue.ToUpper() ?? "") == (rng2CellValue.ToUpper() ?? "")))
                            {

                                this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)this.CBFillBackground.BackColor;

                                this.count = ((this.count) + (1));
                                coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));
                            }

nextLoop6:
                            ;

                        }
                    }
                }

                // 1st Range >> 2nd Range >> radbtnSameValues checked >> case sensitive unchecked >> only font color is selected >> OK
                else if ((((this.checkBoxFillBack.Checked) == (false)) & ((this.checkBoxFillFont.Checked) == (true))))
                {
                    var loopTo14 = this.firstInputRng.Rows.Count;
                    for (i = 1; i <= loopTo14; i++)
                    {
                        var loopTo15 = this.firstInputRng.Columns.Count;
                        for (j = 1; j <= loopTo15; j++)
                        {
                            rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                            rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                            if ((rng1CellValue is null & rng2CellValue is null))
                            {

                                this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)this.CbFillFont.BackColor;

                                this.count = ((this.count) + (1));
                                coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));

                                goto nextLoop7;
                            }

                            else if ((rng1CellValue is null | rng2CellValue is null))
                            {
                                goto nextLoop7;

                            }


                            if ((global::Microsoft.VisualBasic.Information.VarType(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value) != global::Microsoft.VisualBasic.Information.VarType(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                            {
                                goto nextLoop7;
                            }

                            else if (((rng1CellValue.ToUpper() ?? "") == (rng2CellValue.ToUpper() ?? "")))
                            {

                                this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)this.CbFillFont.BackColor;

                                this.count = ((this.count) + (1));
                                coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));
                            }

nextLoop7:
                            ;

                        }
                    }
                }


                else
                {
                    // 1st Range >> 2nd Range >> radbtnSameValues checked >> case sensitive unchecked >> fill/font color not selected >> OK
                    var loopTo10 = this.firstInputRng.Rows.Count;
                    for (i = 1; i <= loopTo10; i++)
                    {
                        var loopTo11 = this.firstInputRng.Columns.Count;
                        for (j = 1; j <= loopTo11; j++)
                        {
                            rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                            rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                            if ((rng1CellValue is null & rng2CellValue is null))
                            {

                                this.count = ((this.count) + (1));
                                coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));

                                goto nextLoop8;
                            }

                            else if ((rng1CellValue is null | rng2CellValue is null))
                            {
                                goto nextLoop8;

                            }


                            if ((global::Microsoft.VisualBasic.Information.VarType(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value) != global::Microsoft.VisualBasic.Information.VarType(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                            {
                                goto nextLoop8;
                            }

                            else if (((rng1CellValue.ToUpper() ?? "") == (rng2CellValue.ToUpper() ?? "")))
                            {
                                this.count = ((this.count) + (1));
                                coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));
                            }

nextLoop8:
                            ;

                        }
                    }


                }
            }

            else if (((this.radBtnDifferentValues.Checked) == (true)))
            {
                if (((this.checkBoxCase.Checked) == (true)))
                {

                    // 1st Range >> 2nd Range >> radBtnDifferentValues checked >> case sensitive checked >> fill/font color both are selected >> OK
                    if ((((this.checkBoxFillBack.Checked) == (true)) & ((this.checkBoxFillFont.Checked) == (true))))
                    {
                        var loopTo16 = this.firstInputRng.Rows.Count;
                        for (i = 1; i <= loopTo16; i++)
                        {
                            var loopTo17 = this.firstInputRng.Columns.Count;
                            for (j = 1; j <= loopTo17; j++)
                            {

                                rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                                if (((((rng1CellValue is null && (rng2CellValue is not null)))) || ((((rng1CellValue is not null) && rng2CellValue is null)))))
                                {
                                    goto nextLoop9;
                                }

                                if ((global::Microsoft.VisualBasic.Information.VarType(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value) != global::Microsoft.VisualBasic.Information.VarType(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                {
                                    goto nextLoop9;
                                }

                                else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value, this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value, false)))
                                {
nextLoop9:
                                    ;

                                    this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)this.CBFillBackground.BackColor;
                                    this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)this.CbFillFont.BackColor;
                                    this.count = ((this.count) + (1));
                                    coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));

                                }

                            }
                        }
                    }

                    // 1st Range >> 2nd Range >> radBtnDifferentValues checked >> case sensitive checked >> only fill color is selected >> OK
                    else if ((((this.checkBoxFillBack.Checked) == (true)) & ((this.checkBoxFillFont.Checked) == (false))))
                    {
                        var loopTo20 = this.firstInputRng.Rows.Count;
                        for (i = 1; i <= loopTo20; i++)
                        {
                            var loopTo21 = this.firstInputRng.Columns.Count;
                            for (j = 1; j <= loopTo21; j++)
                            {

                                rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                                if (((((rng1CellValue is null && (rng2CellValue is not null)))) || ((((rng1CellValue is not null) && rng2CellValue is null)))))
                                {
                                    goto nextLoop10;
                                }

                                if ((global::Microsoft.VisualBasic.Information.VarType(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value) != global::Microsoft.VisualBasic.Information.VarType(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                {
                                    goto nextLoop10;
                                }

                                else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value, this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value, false)))
                                {
nextLoop10:
                                    ;

                                    this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)this.CBFillBackground.BackColor;

                                    this.count = ((this.count) + (1));
                                    coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));

                                }

                            }
                        }
                    }

                    // 1st Range >> 2nd Range >> radBtnDifferentValues checked >> case sensitive checked >> only font color is selected >> OK
                    else if ((((this.checkBoxFillBack.Checked) == (false)) & ((this.checkBoxFillFont.Checked) == (true))))
                    {
                        var loopTo22 = this.firstInputRng.Rows.Count;
                        for (i = 1; i <= loopTo22; i++)
                        {
                            var loopTo23 = this.firstInputRng.Columns.Count;
                            for (j = 1; j <= loopTo23; j++)
                            {

                                rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                                if (((((rng1CellValue is null && (rng2CellValue is not null)))) || ((((rng1CellValue is not null) && rng2CellValue is null)))))
                                {
                                    goto nextLoop11;
                                }

                                if ((global::Microsoft.VisualBasic.Information.VarType(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value) != global::Microsoft.VisualBasic.Information.VarType(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                {
                                    goto nextLoop11;
                                }

                                else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value, this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value, false)))
                                {
nextLoop11:
                                    ;

                                    this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)this.CbFillFont.BackColor;

                                    this.count = ((this.count) + (1));
                                    coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));

                                }
                            }
                        }
                    }
                    else
                    {
                        // 1st Range >> 2nd Range >> radBtnDifferentValues checked >> case sensitive checked >> fill/font color not selected >> OK

                        var loopTo18 = this.firstInputRng.Rows.Count;
                        for (i = 1; i <= loopTo18; i++)
                        {
                            var loopTo19 = this.firstInputRng.Columns.Count;
                            for (j = 1; j <= loopTo19; j++)
                            {

                                rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                                if (((((rng1CellValue is null && (rng2CellValue is not null)))) || ((((rng1CellValue is not null) && rng2CellValue is null)))))
                                {
                                    goto nextLoop12;
                                }

                                if ((global::Microsoft.VisualBasic.Information.VarType(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value) != global::Microsoft.VisualBasic.Information.VarType(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                {
                                    goto nextLoop12;
                                }

                                else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value, this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value, false)))
                                {
nextLoop12:
                                    ;

                                    this.count = ((this.count) + (1));
                                    coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));

                                }

                            }
                        }


                    }
                }

                // 1st Range >> 2nd Range >> radBtnDifferentValues checked >> case sensitive unchecked >> fill/font color both selected >> OK
                else if ((((this.checkBoxFillBack.Checked) == (true)) & ((this.checkBoxFillFont.Checked) == (true))))
                {
                    var loopTo24 = this.firstInputRng.Rows.Count;
                    for (i = 1; i <= loopTo24; i++)
                    {
                        var loopTo25 = this.firstInputRng.Columns.Count;
                        for (j = 1; j <= loopTo25; j++)
                        {
                            rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                            rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                            if (((((rng1CellValue is null && (rng2CellValue is not null)))) || ((((rng1CellValue is not null) && rng2CellValue is null)))))
                            {
                                goto nextLoop13;
                            }

                            if ((global::Microsoft.VisualBasic.Information.VarType(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value) != global::Microsoft.VisualBasic.Information.VarType(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                            {
                                goto nextLoop13;
                            }

                            else if (((rng1CellValue.ToUpper() ?? "") != (rng2CellValue.ToUpper() ?? "")))
                            {
nextLoop13:
                                ;

                                this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)this.CBFillBackground.BackColor;

                                this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)this.CbFillFont.BackColor;
                                this.count = ((this.count) + (1));
                                coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));

                            }

                        }
                    }
                }

                // 1st Range >> 2nd Range >> radBtnDifferentValues checked >> case sensitive unchecked >> only fill color is selected >> OK
                else if ((((this.checkBoxFillBack.Checked) == (true)) & ((this.checkBoxFillFont.Checked) == (false))))
                {
                    var loopTo28 = this.firstInputRng.Rows.Count;
                    for (i = 1; i <= loopTo28; i++)
                    {
                        var loopTo29 = this.firstInputRng.Columns.Count;
                        for (j = 1; j <= loopTo29; j++)
                        {
                            rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                            rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                            if (((((rng1CellValue is null && (rng2CellValue is not null)))) || ((((rng1CellValue is not null) && rng2CellValue is null)))))
                            {
                                goto nextLoop14;
                            }

                            if ((global::Microsoft.VisualBasic.Information.VarType(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value) != global::Microsoft.VisualBasic.Information.VarType(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                            {
                                goto nextLoop14;
                            }

                            else if (((rng1CellValue.ToUpper() ?? "") != (rng2CellValue.ToUpper() ?? "")))
                            {
nextLoop14:
                                ;

                                this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)this.CBFillBackground.BackColor;

                                this.count = ((this.count) + (1));
                                coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));

                            }

                        }
                    }
                }

                // 1st Range >> 2nd Range >> radBtnDifferentValues checked >> case sensitive unchecked >> only font color is selected >> OK
                else if ((((this.checkBoxFillBack.Checked) == (false)) & ((this.checkBoxFillFont.Checked) == (true))))
                {
                    var loopTo30 = this.firstInputRng.Rows.Count;
                    for (i = 1; i <= loopTo30; i++)
                    {
                        var loopTo31 = this.firstInputRng.Columns.Count;
                        for (j = 1; j <= loopTo31; j++)
                        {
                            rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                            rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                            if (((((rng1CellValue is null && (rng2CellValue is not null)))) || ((((rng1CellValue is not null) && rng2CellValue is null)))))
                            {
                                goto nextLoop15;
                            }

                            if ((global::Microsoft.VisualBasic.Information.VarType(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value) != global::Microsoft.VisualBasic.Information.VarType(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                            {
                                goto nextLoop15;
                            }

                            else if (((rng1CellValue.ToUpper() ?? "") != (rng2CellValue.ToUpper() ?? "")))
                            {
nextLoop15:
                                ;

                                this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)this.CbFillFont.BackColor;
                                this.count = ((this.count) + (1));
                                coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));

                            }

                        }

                    }
                }



                else
                {
                    // 1st Range >> 2nd Range >> radBtnDifferentValues checked >> case sensitive unchecked >> fill/font color not selected >> OK
                    var loopTo26 = this.firstInputRng.Rows.Count;
                    for (i = 1; i <= loopTo26; i++)
                    {
                        var loopTo27 = this.firstInputRng.Columns.Count;
                        for (j = 1; j <= loopTo27; j++)
                        {
                            rng1CellValue = Conversions.ToString(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                            rng2CellValue = Conversions.ToString(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value);

                            if ((rng1CellValue is null & rng2CellValue is null))
                            {
                                continue;
                            }

                            else if (((((rng1CellValue is null && (rng2CellValue is not null)))) || ((((rng1CellValue is not null) && rng2CellValue is null)))))
                            {
                                goto nextLoop16;
                            }

                            if ((global::Microsoft.VisualBasic.Information.VarType(this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value) != global::Microsoft.VisualBasic.Information.VarType(this.secondInputRng.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                            {
                                goto nextLoop16;
                            }

                            else if (((rng1CellValue.ToUpper() ?? "") != (rng2CellValue.ToUpper() ?? "")))
                            {
nextLoop16:
                                ;

                                this.count = ((this.count) + (1));
                                coloredRng = Conversions.ToString(Operators.ConcatenateObject((coloredRng + ","), this.firstInputRng.Cells[(global::System.Object)i, (global::System.Object)j].address));

                            }

                        }
                    }

                }

            }

            this.Dispose();


            this.firstInputRng.Worksheet.Activate();

            global::Microsoft.VisualBasic.Interaction.MsgBox(this.count + " cell(s) found.", global::Microsoft.VisualBasic.MsgBoxStyle.Information, "SOFTEKO");
            if (string.IsNullOrEmpty(coloredRng))
            {
                return;
            }
            else
            {
                coloredRng = global::Microsoft.VisualBasic.Strings.Right(coloredRng, (global::Microsoft.VisualBasic.Strings.Len(coloredRng)) - (1));
                this.firstInputRng.Worksheet.get_Range(coloredRng).Select();
            }



        }


        public void Display()
        {

            try
            {

                this.CP_Input_Range1.Controls.Clear();
                this.CP_Input_Range2.Controls.Clear();
                this.CP_Output_Range.Controls.Clear();


                global::Microsoft.Office.Interop.Excel.Range displayRng;
                global::System.Int64 lblColor;
                global::System.Drawing.Color rgbColor;

                if ((string.IsNullOrEmpty(this.txtSourceRange1.Text) | this.firstInputRng is null))
                {
                    this.CP_Input_Range1.Controls.Clear();
                    goto secondDisplay;
                }


                if (((this.firstInputRng.Rows.Count) > (50)))
                {
                    displayRng = (global::Microsoft.Office.Interop.Excel.Range)this.firstInputRng.Rows["1:50"];
                }
                else
                {
                    displayRng = this.firstInputRng;
                }


                var height = default(global::System.Double);
                var width = default(global::System.Double);

                if (((displayRng.Rows.Count) <= (4)))
                {
                    height = ((global::System.Double)(this.CP_Input_Range1.Height) / (global::System.Double)(displayRng.Rows.Count));
                }
                else
                {
                    height = (((119d) / (4d)));
                }

                if (((displayRng.Columns.Count) <= (3)))
                {
                    width = ((global::System.Double)(this.CP_Input_Range1.Width) / (global::System.Double)(displayRng.Columns.Count));
                }
                else
                {
                    width = (((260d) / (3d)));
                }

                for (global::System.Int32 i = 1, loopTo = displayRng.Rows.Count; i <= loopTo; i++)
                {
                    for (global::System.Int32 j = 1, loopTo1 = displayRng.Columns.Count; j <= loopTo1; j++)
                    {
                        var label = new global::System.Windows.Forms.Label();
                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                        label.Height = (global::System.Int32)Math.Round(height);
                        label.Width = (global::System.Int32)Math.Round(width);
                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                        if (((this.checkBoxFormatting.Checked) == (true)))
                        {

                            // background fill color
                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                            label.BackColor = rgbColor;

                            // font color
                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                            label.ForeColor = rgbColor;

                        }
                        // label.BackColor = displayRng.Cells(i, j).interior.color
                        // label.ForeColor = displayRng.Cells(i, j).font.color

                        this.CP_Input_Range1.Controls.Add(label);
                    }
                }

                this.CP_Input_Range1.AutoScroll = true;

secondDisplay:
                ;


                if ((string.IsNullOrEmpty(this.txtSourceRange2.Text) | this.secondInputRng is null))
                {
                    this.CP_Input_Range2.Controls.Clear();
                    return;
                }

                global::Microsoft.Office.Interop.Excel.Range displayRng2;
                if (((this.secondInputRng.Rows.Count) > (50)))
                {
                    displayRng2 = (global::Microsoft.Office.Interop.Excel.Range)this.secondInputRng.Rows["1:50"];
                }
                else
                {
                    displayRng2 = this.secondInputRng;
                }


                global::System.Double height2;
                global::System.Double width2;

                if (((displayRng2.Rows.Count) <= (4)))
                {
                    height2 = ((global::System.Double)(this.CP_Input_Range2.Height) / (global::System.Double)(displayRng2.Rows.Count));
                }
                else
                {
                    height2 = (((119d) / (4d)));
                }

                if (((displayRng2.Columns.Count) <= (3)))
                {
                    width2 = ((global::System.Double)(this.CP_Input_Range2.Width) / (global::System.Double)(displayRng2.Columns.Count));
                }
                else
                {
                    width2 = (((260d) / (3d)));
                }

                for (global::System.Int32 i = 1, loopTo2 = displayRng2.Rows.Count; i <= loopTo2; i++)
                {
                    for (global::System.Int32 j = 1, loopTo3 = displayRng2.Columns.Count; j <= loopTo3; j++)
                    {
                        var label = new global::System.Windows.Forms.Label();
                        label.Text = Conversions.ToString(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width2)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height2)));
                        label.Height = (global::System.Int32)Math.Round(height2);
                        label.Width = (global::System.Int32)Math.Round(width2);
                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                        if (((this.checkBoxFormatting.Checked) == (true)))
                        {

                            // backgroud fill color
                            lblColor = Conversions.ToLong(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                            label.BackColor = rgbColor;

                            // font color
                            lblColor = Conversions.ToLong(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                            label.ForeColor = rgbColor;

                        }



                        this.CP_Input_Range2.Controls.Add(label);
                    }
                }

                this.CP_Input_Range2.AutoScroll = true;

                if ((string.IsNullOrEmpty(this.txtSourceRange1.Text) | this.firstInputRng is null))
                {
                    return;
                }

                if (((this.firstInputRng.Rows.Count) > (50)))
                {
                    displayRng = (global::Microsoft.Office.Interop.Excel.Range)this.firstInputRng.Rows["1:50"];
                }
                else
                {
                    displayRng = this.firstInputRng;
                }

                if ((((displayRng.Rows.Count) != (displayRng2.Rows.Count)) | ((displayRng.Columns.Count) != (displayRng2.Columns.Count))))
                {
                    return;
                }


                if (((this.radBtnSameValues.Checked) == (true)))
                {

                    if (((this.checkBoxCase.Checked) == (true)))
                    {

                        // 1st range >> 2nd range >> radBtnSameValues checked >> case sensitive checked >> both fill/font color selected
                        if ((((this.checkBoxFillBack.Checked) == (true)) & ((this.checkBoxFillFont.Checked) == (true))))
                        {
                            for (global::System.Int32 i = 1, loopTo4 = displayRng.Rows.Count; i <= loopTo4; i++)
                            {
                                for (global::System.Int32 j = 1, loopTo5 = displayRng.Columns.Count; j <= loopTo5; j++)
                                {

                                    if ((global::Microsoft.VisualBasic.Information.VarType(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value) == global::Microsoft.VisualBasic.Information.VarType(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                    {
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value, displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value, false)))
                                        {

                                            var label = new global::System.Windows.Forms.Label();
                                            label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                            label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                            label.Height = (global::System.Int32)Math.Round(height);
                                            label.Width = (global::System.Int32)Math.Round(width);
                                            label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                            label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                            label.BackColor = this.CBFillBackground.BackColor;
                                            label.ForeColor = this.CbFillFont.BackColor;

                                            this.CP_Output_Range.Controls.Add(label);
                                        }
                                        else
                                        {
                                            var label = new global::System.Windows.Forms.Label();
                                            label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                            label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                            label.Height = (global::System.Int32)Math.Round(height);
                                            label.Width = (global::System.Int32)Math.Round(width);
                                            label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                            label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                            // label.BackColor = Color.Transparent
                                            // label.ForeColor = Nothing
                                            if (((this.checkBoxFormatting.Checked) == (true)))
                                            {

                                                // background fill color
                                                lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                                rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                                label.BackColor = rgbColor;

                                                // font color
                                                lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                                rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                                label.ForeColor = rgbColor;
                                            }

                                            else
                                            {
                                                label.BackColor = global::System.Drawing.Color.Transparent;
                                                label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                            }

                                            this.CP_Output_Range.Controls.Add(label);

                                        }
                                    }

                                    else
                                    {
                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                        // label.BackColor = Color.Transparent
                                        // label.ForeColor = Nothing


                                        if (((this.checkBoxFormatting.Checked) == (true)))
                                        {

                                            // background fill color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.BackColor = rgbColor;

                                            // font color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.ForeColor = rgbColor;
                                        }

                                        else
                                        {
                                            label.BackColor = global::System.Drawing.Color.Transparent;
                                            label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                        }

                                        this.CP_Output_Range.Controls.Add(label);
                                    }
                                }
                            }
                        }

                        // 1st range >> 2nd range >> radBtnSameValues checked >> case sensitive checked >> only fill color is selected
                        else if ((((this.checkBoxFillBack.Checked) == (true)) & ((this.checkBoxFillFont.Checked) == (false))))
                        {
                            for (global::System.Int32 i = 1, loopTo8 = displayRng.Rows.Count; i <= loopTo8; i++)
                            {
                                for (global::System.Int32 j = 1, loopTo9 = displayRng.Columns.Count; j <= loopTo9; j++)
                                {

                                    if ((global::Microsoft.VisualBasic.Information.VarType(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value) == global::Microsoft.VisualBasic.Information.VarType(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                    {
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value, displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value, false)))
                                        {

                                            var label = new global::System.Windows.Forms.Label();
                                            label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                            label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                            label.Height = (global::System.Int32)Math.Round(height);
                                            label.Width = (global::System.Int32)Math.Round(width);
                                            label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                            label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                            label.BackColor = this.CBFillBackground.BackColor;

                                            if (((this.checkBoxFormatting.Checked) == (true)))
                                            {
                                                lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                                rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                                label.ForeColor = rgbColor;
                                            }
                                            else
                                            {
                                                label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));
                                            }


                                            this.CP_Output_Range.Controls.Add(label);
                                        }
                                        else
                                        {
                                            var label = new global::System.Windows.Forms.Label();
                                            label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                            label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                            label.Height = (global::System.Int32)Math.Round(height);
                                            label.Width = (global::System.Int32)Math.Round(width);
                                            label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                            label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                            // label.BackColor = Color.Transparent
                                            // label.ForeColor = Nothing

                                            if (((this.checkBoxFormatting.Checked) == (true)))
                                            {

                                                // background fill color
                                                lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                                rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                                label.BackColor = rgbColor;

                                                // font color
                                                lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                                rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                                label.ForeColor = rgbColor;
                                            }

                                            else
                                            {
                                                label.BackColor = global::System.Drawing.Color.Transparent;
                                                label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                            }

                                            this.CP_Output_Range.Controls.Add(label);

                                        }
                                    }

                                    else
                                    {
                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                        // label.BackColor = Color.Transparent
                                        // label.ForeColor = Nothing


                                        if (((this.checkBoxFormatting.Checked) == (true)))
                                        {

                                            // background fill color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.BackColor = rgbColor;

                                            // font color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.ForeColor = rgbColor;
                                        }

                                        else
                                        {
                                            label.BackColor = global::System.Drawing.Color.Transparent;
                                            label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                        }

                                        this.CP_Output_Range.Controls.Add(label);

                                    }
                                }
                            }
                        }

                        // 1st range >> 2nd range >> radBtnSameValues checked >> case sensitive checked >> only font color is selected
                        else if ((((this.checkBoxFillBack.Checked) == (false)) & ((this.checkBoxFillFont.Checked) == (true))))
                        {
                            for (global::System.Int32 i = 1, loopTo10 = displayRng.Rows.Count; i <= loopTo10; i++)
                            {
                                for (global::System.Int32 j = 1, loopTo11 = displayRng.Columns.Count; j <= loopTo11; j++)
                                {

                                    if ((global::Microsoft.VisualBasic.Information.VarType(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value) == global::Microsoft.VisualBasic.Information.VarType(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                    {
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value, displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value, false)))
                                        {

                                            var label = new global::System.Windows.Forms.Label();
                                            label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                            label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                            label.Height = (global::System.Int32)Math.Round(height);
                                            label.Width = (global::System.Int32)Math.Round(width);
                                            label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                            label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                            label.ForeColor = this.CbFillFont.BackColor;

                                            if (((this.checkBoxFormatting.Checked) == (true)))
                                            {
                                                lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                                rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                                label.BackColor = rgbColor;
                                            }
                                            else
                                            {
                                                label.BackColor = global::System.Drawing.Color.Transparent;

                                            }


                                            this.CP_Output_Range.Controls.Add(label);
                                        }
                                        else
                                        {
                                            var label = new global::System.Windows.Forms.Label();
                                            label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                            label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                            label.Height = (global::System.Int32)Math.Round(height);
                                            label.Width = (global::System.Int32)Math.Round(width);
                                            label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                            label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                            if (((this.checkBoxFormatting.Checked) == (true)))
                                            {

                                                // background fill color
                                                lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                                rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                                label.BackColor = rgbColor;

                                                // font color
                                                lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                                rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                                label.ForeColor = rgbColor;
                                            }

                                            else
                                            {
                                                label.BackColor = global::System.Drawing.Color.Transparent;
                                                label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                            }

                                            this.CP_Output_Range.Controls.Add(label);

                                        }
                                    }

                                    else
                                    {
                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                        if (((this.checkBoxFormatting.Checked) == (true)))
                                        {

                                            // background fill color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.BackColor = rgbColor;

                                            // font color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.ForeColor = rgbColor;
                                        }

                                        else
                                        {
                                            label.BackColor = global::System.Drawing.Color.Transparent;
                                            label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                        }

                                        this.CP_Output_Range.Controls.Add(label);

                                    }
                                }
                            }
                        }

                        else
                        {
                            // 1st range >> 2nd range >> radBtnSameValues checked >> case sensitive checked >> fill/font color not selected
                            for (global::System.Int32 i = 1, loopTo6 = displayRng.Rows.Count; i <= loopTo6; i++)
                            {
                                for (global::System.Int32 j = 1, loopTo7 = displayRng.Columns.Count; j <= loopTo7; j++)
                                {
                                    var label = new global::System.Windows.Forms.Label();
                                    label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                    label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                    label.Height = (global::System.Int32)Math.Round(height);
                                    label.Width = (global::System.Int32)Math.Round(width);
                                    label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                    // label.BackColor = Color.Transparent
                                    // label.ForeColor = Nothing

                                    if (((this.checkBoxFormatting.Checked) == (true)))
                                    {

                                        // background fill color
                                        lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                        rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                        label.BackColor = rgbColor;

                                        // font color
                                        lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                        rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                        label.ForeColor = rgbColor;
                                    }

                                    else
                                    {
                                        label.BackColor = global::System.Drawing.Color.Transparent;
                                        label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                    }

                                    this.CP_Output_Range.Controls.Add(label);

                                }
                            }

                        }
                    }
                    // 1st range >> 2nd range >> radBtnSameValues checked >> case sensitive unchecked >> fill/font color both are selected
                    else if ((((this.checkBoxFillBack.Checked) == (true)) & ((this.checkBoxFillFont.Checked) == (true))))
                    {
                        for (global::System.Int32 i = 1, loopTo12 = displayRng.Rows.Count; i <= loopTo12; i++)
                        {
                            for (global::System.Int32 j = 1, loopTo13 = displayRng.Columns.Count; j <= loopTo13; j++)
                            {
                                this.rng1CellValue = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                this.rng2CellValue = Conversions.ToString(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value);

                                if ((global::Microsoft.VisualBasic.Information.VarType(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value) == global::Microsoft.VisualBasic.Information.VarType(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                {

                                    if (((this.rng1CellValue.ToUpper() ?? "") == (this.rng2CellValue.ToUpper() ?? "")))
                                    {

                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                        label.BackColor = this.CBFillBackground.BackColor;
                                        label.ForeColor = this.CbFillFont.BackColor;

                                        this.CP_Output_Range.Controls.Add(label);
                                    }
                                    else
                                    {
                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                        // label.BackColor = Color.Transparent
                                        // label.ForeColor = Nothing

                                        if (((this.checkBoxFormatting.Checked) == (true)))
                                        {

                                            // background fill color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.BackColor = rgbColor;

                                            // font color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.ForeColor = rgbColor;
                                        }

                                        else
                                        {
                                            label.BackColor = global::System.Drawing.Color.Transparent;
                                            label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                        }

                                        this.CP_Output_Range.Controls.Add(label);

                                    }
                                }

                                else
                                {
                                    var label = new global::System.Windows.Forms.Label();
                                    label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                    label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                    label.Height = (global::System.Int32)Math.Round(height);
                                    label.Width = (global::System.Int32)Math.Round(width);
                                    label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                    if (((this.checkBoxFormatting.Checked) == (true)))
                                    {

                                        // background fill color
                                        lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                        rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                        label.BackColor = rgbColor;

                                        // font color
                                        lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                        rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                        label.ForeColor = rgbColor;
                                    }

                                    else
                                    {
                                        label.BackColor = global::System.Drawing.Color.Transparent;
                                        label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                    }

                                    this.CP_Output_Range.Controls.Add(label);

                                }
                            }
                        }
                    }

                    // 1st range >> 2nd range >> radBtnSameValues checked >> case sensitive unchecked >> only fill color is selected
                    else if ((((this.checkBoxFillBack.Checked) == (true)) & ((this.checkBoxFillFont.Checked) == (false))))
                    {
                        for (global::System.Int32 i = 1, loopTo16 = displayRng.Rows.Count; i <= loopTo16; i++)
                        {
                            for (global::System.Int32 j = 1, loopTo17 = displayRng.Columns.Count; j <= loopTo17; j++)
                            {
                                this.rng1CellValue = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                this.rng2CellValue = Conversions.ToString(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value);

                                if ((global::Microsoft.VisualBasic.Information.VarType(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value) == global::Microsoft.VisualBasic.Information.VarType(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                {

                                    if (((this.rng1CellValue.ToUpper() ?? "") == (this.rng2CellValue.ToUpper() ?? "")))
                                    {

                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                        label.BackColor = this.CBFillBackground.BackColor;

                                        if (((this.checkBoxFormatting.Checked) == (true)))
                                        {
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.ForeColor = rgbColor;
                                        }
                                        else
                                        {
                                            label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                        }

                                        this.CP_Output_Range.Controls.Add(label);
                                    }
                                    else
                                    {
                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                        if (((this.checkBoxFormatting.Checked) == (true)))
                                        {

                                            // background fill color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.BackColor = rgbColor;

                                            // font color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.ForeColor = rgbColor;
                                        }

                                        else
                                        {
                                            label.BackColor = global::System.Drawing.Color.Transparent;
                                            label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                        }

                                        this.CP_Output_Range.Controls.Add(label);

                                    }
                                }

                                else
                                {
                                    var label = new global::System.Windows.Forms.Label();
                                    label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                    label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                    label.Height = (global::System.Int32)Math.Round(height);
                                    label.Width = (global::System.Int32)Math.Round(width);
                                    label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                    if (((this.checkBoxFormatting.Checked) == (true)))
                                    {

                                        // background fill color
                                        lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                        rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                        label.BackColor = rgbColor;

                                        // font color
                                        lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                        rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                        label.ForeColor = rgbColor;
                                    }

                                    else
                                    {
                                        label.BackColor = global::System.Drawing.Color.Transparent;
                                        label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                    }


                                    this.CP_Output_Range.Controls.Add(label);

                                }
                            }
                        }
                    }

                    // 1st range >> 2nd range >> radBtnSameValues checked >> case sensitive unchecked >> only font color is selected
                    else if ((((this.checkBoxFillBack.Checked) == (false)) & ((this.checkBoxFillFont.Checked) == (true))))
                    {
                        for (global::System.Int32 i = 1, loopTo18 = displayRng.Rows.Count; i <= loopTo18; i++)
                        {
                            for (global::System.Int32 j = 1, loopTo19 = displayRng.Columns.Count; j <= loopTo19; j++)
                            {
                                this.rng1CellValue = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                this.rng2CellValue = Conversions.ToString(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value);


                                if ((global::Microsoft.VisualBasic.Information.VarType(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value) == global::Microsoft.VisualBasic.Information.VarType(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                {

                                    if (((this.rng1CellValue.ToUpper() ?? "") == (this.rng2CellValue.ToUpper() ?? "")))
                                    {

                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                        label.ForeColor = this.CbFillFont.BackColor;

                                        if (((this.checkBoxFormatting.Checked) == (true)))
                                        {
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.BackColor = rgbColor;
                                        }
                                        else
                                        {
                                            label.BackColor = global::System.Drawing.Color.Transparent;
                                        }


                                        this.CP_Output_Range.Controls.Add(label);
                                    }
                                    else
                                    {
                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                        if (((this.checkBoxFormatting.Checked) == (true)))
                                        {

                                            // background fill color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.BackColor = rgbColor;

                                            // font color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.ForeColor = rgbColor;
                                        }

                                        else
                                        {
                                            label.BackColor = global::System.Drawing.Color.Transparent;
                                            label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                        }

                                        this.CP_Output_Range.Controls.Add(label);

                                    }
                                }

                                else
                                {
                                    var label = new global::System.Windows.Forms.Label();
                                    label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                    label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                    label.Height = (global::System.Int32)Math.Round(height);
                                    label.Width = (global::System.Int32)Math.Round(width);
                                    label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                    if (((this.checkBoxFormatting.Checked) == (true)))
                                    {

                                        // background fill color
                                        lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                        rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                        label.BackColor = rgbColor;

                                        // font color
                                        lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                        rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                        label.ForeColor = rgbColor;
                                    }

                                    else
                                    {
                                        label.BackColor = global::System.Drawing.Color.Transparent;
                                        label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                    }

                                    this.CP_Output_Range.Controls.Add(label);

                                }
                            }
                        }
                    }

                    // 1st range >> 2nd range >> radBtnSameValues checked >> case sensitive unchecked >> fill/font color not selected
                    else
                    {
                        for (global::System.Int32 i = 1, loopTo14 = displayRng.Rows.Count; i <= loopTo14; i++)
                        {
                            for (global::System.Int32 j = 1, loopTo15 = displayRng.Columns.Count; j <= loopTo15; j++)
                            {
                                this.rng1CellValue = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                this.rng2CellValue = Conversions.ToString(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value);

                                var label = new global::System.Windows.Forms.Label();
                                label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                label.Height = (global::System.Int32)Math.Round(height);
                                label.Width = (global::System.Int32)Math.Round(width);
                                label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                if (((this.checkBoxFormatting.Checked) == (true)))
                                {

                                    // background fill color
                                    lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                    rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                    label.BackColor = rgbColor;

                                    // font color
                                    lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                    rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                    label.ForeColor = rgbColor;
                                }

                                else
                                {
                                    label.BackColor = global::System.Drawing.Color.Transparent;
                                    label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                }

                                this.CP_Output_Range.Controls.Add(label);


                            }
                        }


                    }
                }


                else if (((this.radBtnDifferentValues.Checked) == (true)))
                {

                    if (((this.checkBoxCase.Checked) == (true)))
                    {

                        // 1st range >> 2nd range >> radBtnDifferentValues checked >> case sensitive checked >> fill/font color both are selected
                        if ((((this.checkBoxFillBack.Checked) == (true)) & ((this.checkBoxFillFont.Checked) == (true))))
                        {
                            for (global::System.Int32 i = 1, loopTo20 = displayRng.Rows.Count; i <= loopTo20; i++)
                            {
                                for (global::System.Int32 j = 1, loopTo21 = displayRng.Columns.Count; j <= loopTo21; j++)
                                {

                                    if ((global::Microsoft.VisualBasic.Information.VarType(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value) == global::Microsoft.VisualBasic.Information.VarType(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                    {
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value, displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value, false)))
                                        {

                                            var label = new global::System.Windows.Forms.Label();
                                            label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                            label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                            label.Height = (global::System.Int32)Math.Round(height);
                                            label.Width = (global::System.Int32)Math.Round(width);
                                            label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                            label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                            label.BackColor = this.CBFillBackground.BackColor;
                                            label.ForeColor = this.CbFillFont.BackColor;

                                            this.CP_Output_Range.Controls.Add(label);
                                        }
                                        else
                                        {
                                            var label = new global::System.Windows.Forms.Label();
                                            label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                            label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                            label.Height = (global::System.Int32)Math.Round(height);
                                            label.Width = (global::System.Int32)Math.Round(width);
                                            label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                            label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                            if (((this.checkBoxFormatting.Checked) == (true)))
                                            {

                                                // background fill color
                                                lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                                rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                                label.BackColor = rgbColor;

                                                // font color
                                                lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                                rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                                label.ForeColor = rgbColor;
                                            }

                                            else
                                            {
                                                label.BackColor = global::System.Drawing.Color.Transparent;
                                                label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                            }

                                            this.CP_Output_Range.Controls.Add(label);

                                        }
                                    }

                                    else
                                    {
                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                        label.BackColor = this.CBFillBackground.BackColor;
                                        label.ForeColor = this.CbFillFont.BackColor;



                                        this.CP_Output_Range.Controls.Add(label);

                                    }
                                }
                            }
                        }

                        // 1st range >> 2nd range >> radBtnDifferentValues checked >> case sensitive checked >> only fill color is selected
                        else if ((((this.checkBoxFillBack.Checked) == (true)) & ((this.checkBoxFillFont.Checked) == (false))))
                        {
                            for (global::System.Int32 i = 1, loopTo24 = displayRng.Rows.Count; i <= loopTo24; i++)
                            {
                                for (global::System.Int32 j = 1, loopTo25 = displayRng.Columns.Count; j <= loopTo25; j++)
                                {

                                    if ((global::Microsoft.VisualBasic.Information.VarType(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value) == global::Microsoft.VisualBasic.Information.VarType(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                    {
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value, displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value, false)))
                                        {

                                            var label = new global::System.Windows.Forms.Label();
                                            label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                            label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                            label.Height = (global::System.Int32)Math.Round(height);
                                            label.Width = (global::System.Int32)Math.Round(width);
                                            label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                            label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                            label.BackColor = this.CBFillBackground.BackColor;

                                            if (((this.checkBoxFormatting.Checked) == (true)))
                                            {
                                                lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                                rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                                label.ForeColor = rgbColor;
                                            }
                                            else
                                            {
                                                label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));
                                            }


                                            this.CP_Output_Range.Controls.Add(label);
                                        }
                                        else
                                        {
                                            var label = new global::System.Windows.Forms.Label();
                                            label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                            label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                            label.Height = (global::System.Int32)Math.Round(height);
                                            label.Width = (global::System.Int32)Math.Round(width);
                                            label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                            label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                            // label.BackColor = Color.Transparent
                                            // label.ForeColor = Nothing

                                            if (((this.checkBoxFormatting.Checked) == (true)))
                                            {

                                                // background fill color
                                                lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                                rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                                label.BackColor = rgbColor;

                                                // font color
                                                lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                                rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                                label.ForeColor = rgbColor;
                                            }

                                            else
                                            {
                                                label.BackColor = global::System.Drawing.Color.Transparent;
                                                label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                            }

                                            this.CP_Output_Range.Controls.Add(label);

                                        }
                                    }

                                    else
                                    {
                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                        label.BackColor = this.CBFillBackground.BackColor;
                                        label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));


                                        this.CP_Output_Range.Controls.Add(label);

                                    }
                                }
                            }
                        }

                        // 1st range >> 2nd range >> radBtnDifferentValues checked >> case sensitive checked >> only font color is selected
                        else if ((((this.checkBoxFillBack.Checked) == (false)) & ((this.checkBoxFillFont.Checked) == (true))))
                        {
                            for (global::System.Int32 i = 1, loopTo26 = displayRng.Rows.Count; i <= loopTo26; i++)
                            {
                                for (global::System.Int32 j = 1, loopTo27 = displayRng.Columns.Count; j <= loopTo27; j++)
                                {

                                    if ((global::Microsoft.VisualBasic.Information.VarType(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value) == global::Microsoft.VisualBasic.Information.VarType(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                    {
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value, displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value, false)))
                                        {

                                            var label = new global::System.Windows.Forms.Label();
                                            label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                            label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                            label.Height = (global::System.Int32)Math.Round(height);
                                            label.Width = (global::System.Int32)Math.Round(width);
                                            label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                            label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                            label.ForeColor = this.CbFillFont.BackColor;
                                            if (((this.checkBoxFormatting.Checked) == (true)))
                                            {

                                                lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                                rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                                label.BackColor = rgbColor;
                                            }
                                            else
                                            {
                                                label.BackColor = global::System.Drawing.Color.Transparent;
                                            }

                                            this.CP_Output_Range.Controls.Add(label);
                                        }
                                        else
                                        {
                                            var label = new global::System.Windows.Forms.Label();
                                            label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                            label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                            label.Height = (global::System.Int32)Math.Round(height);
                                            label.Width = (global::System.Int32)Math.Round(width);
                                            label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                            label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                            if (((this.checkBoxFormatting.Checked) == (true)))
                                            {

                                                // background fill color
                                                lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                                rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                                label.BackColor = rgbColor;

                                                // font color
                                                lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                                rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                                label.ForeColor = rgbColor;
                                            }

                                            else
                                            {
                                                label.BackColor = global::System.Drawing.Color.Transparent;
                                                label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                            }

                                            this.CP_Output_Range.Controls.Add(label);

                                        }
                                    }

                                    else
                                    {
                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                        label.BackColor = global::System.Drawing.Color.Transparent;
                                        label.ForeColor = this.CbFillFont.BackColor;



                                        this.CP_Output_Range.Controls.Add(label);

                                    }
                                }
                            }
                        }

                        // 1st range >> 2nd range >> radBtnDifferentValues checked >> case sensitive checked >> fill/font color not selected
                        else
                        {
                            for (global::System.Int32 i = 1, loopTo22 = displayRng.Rows.Count; i <= loopTo22; i++)
                            {
                                for (global::System.Int32 j = 1, loopTo23 = displayRng.Columns.Count; j <= loopTo23; j++)
                                {

                                    var label = new global::System.Windows.Forms.Label();
                                    label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                    label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                    label.Height = (global::System.Int32)Math.Round(height);
                                    label.Width = (global::System.Int32)Math.Round(width);
                                    label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                    // label.BackColor = Color.Transparent
                                    // label.ForeColor = Nothing

                                    if (((this.checkBoxFormatting.Checked) == (true)))
                                    {

                                        // background fill color
                                        lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                        rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                        label.BackColor = rgbColor;

                                        // font color
                                        lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                        rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                        label.ForeColor = rgbColor;
                                    }

                                    else
                                    {
                                        label.BackColor = global::System.Drawing.Color.Transparent;
                                        label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                    }

                                    this.CP_Output_Range.Controls.Add(label);

                                }
                            }

                        }
                    }





                    // 1st range >> 2nd range >> radBtnDifferentValues checked >> case sensitive unchecked >> fill/font color both are selected
                    else if ((((this.checkBoxFillBack.Checked) == (true)) & ((this.checkBoxFillFont.Checked) == (true))))
                    {
                        for (global::System.Int32 i = 1, loopTo28 = displayRng.Rows.Count; i <= loopTo28; i++)
                        {
                            for (global::System.Int32 j = 1, loopTo29 = displayRng.Columns.Count; j <= loopTo29; j++)
                            {
                                this.rng1CellValue = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                this.rng2CellValue = Conversions.ToString(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value);

                                if ((global::Microsoft.VisualBasic.Information.VarType(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value) == global::Microsoft.VisualBasic.Information.VarType(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                {

                                    if (((this.rng1CellValue.ToUpper() ?? "") != (this.rng2CellValue.ToUpper() ?? "")))
                                    {

                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                        label.BackColor = this.CBFillBackground.BackColor;
                                        label.ForeColor = this.CbFillFont.BackColor;

                                        this.CP_Output_Range.Controls.Add(label);
                                    }
                                    else
                                    {
                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                        if (((this.checkBoxFormatting.Checked) == (true)))
                                        {

                                            // background fill color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.BackColor = rgbColor;

                                            // font color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.ForeColor = rgbColor;
                                        }

                                        else
                                        {
                                            label.BackColor = global::System.Drawing.Color.Transparent;
                                            label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                        }

                                        this.CP_Output_Range.Controls.Add(label);

                                    }
                                }

                                else
                                {
                                    var label = new global::System.Windows.Forms.Label();
                                    label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                    label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                    label.Height = (global::System.Int32)Math.Round(height);
                                    label.Width = (global::System.Int32)Math.Round(width);
                                    label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                    label.BackColor = this.CBFillBackground.BackColor;
                                    label.ForeColor = this.CbFillFont.BackColor;


                                    this.CP_Output_Range.Controls.Add(label);

                                }
                            }
                        }
                    }

                    // 1st range >> 2nd range >> radBtnDifferentValues checked >> case sensitive unchecked >> only fill color is selected
                    else if ((((this.checkBoxFillBack.Checked) == (true)) & ((this.checkBoxFillFont.Checked) == (false))))
                    {
                        for (global::System.Int32 i = 1, loopTo32 = displayRng.Rows.Count; i <= loopTo32; i++)
                        {
                            for (global::System.Int32 j = 1, loopTo33 = displayRng.Columns.Count; j <= loopTo33; j++)
                            {
                                this.rng1CellValue = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                this.rng2CellValue = Conversions.ToString(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value);


                                if ((global::Microsoft.VisualBasic.Information.VarType(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value) == global::Microsoft.VisualBasic.Information.VarType(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                {

                                    if (((this.rng1CellValue.ToUpper() ?? "") != (this.rng2CellValue.ToUpper() ?? "")))
                                    {

                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                        label.BackColor = this.CBFillBackground.BackColor;

                                        if (((this.checkBoxFormatting.Checked) == (true)))
                                        {
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.ForeColor = rgbColor;
                                        }
                                        else
                                        {
                                            label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));
                                        }

                                        this.CP_Output_Range.Controls.Add(label);
                                    }
                                    else
                                    {
                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;


                                        if (((this.checkBoxFormatting.Checked) == (true)))
                                        {

                                            // background fill color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.BackColor = rgbColor;

                                            // font color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.ForeColor = rgbColor;
                                        }

                                        else
                                        {
                                            label.BackColor = global::System.Drawing.Color.Transparent;
                                            label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                        }

                                        this.CP_Output_Range.Controls.Add(label);

                                    }
                                }

                                else
                                {
                                    var label = new global::System.Windows.Forms.Label();
                                    label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                    label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                    label.Height = (global::System.Int32)Math.Round(height);
                                    label.Width = (global::System.Int32)Math.Round(width);
                                    label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                    label.BackColor = this.CBFillBackground.BackColor;
                                    label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));


                                    this.CP_Output_Range.Controls.Add(label);

                                }

                            }
                        }
                    }

                    // 1st range >> 2nd range >> radBtnDifferentValues checked >> case sensitive unchecked >> only font color is selected
                    else if ((((this.checkBoxFillBack.Checked) == (false)) & ((this.checkBoxFillFont.Checked) == (true))))
                    {
                        for (global::System.Int32 i = 1, loopTo34 = displayRng.Rows.Count; i <= loopTo34; i++)
                        {
                            for (global::System.Int32 j = 1, loopTo35 = displayRng.Columns.Count; j <= loopTo35; j++)
                            {
                                this.rng1CellValue = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                this.rng2CellValue = Conversions.ToString(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value);


                                if ((global::Microsoft.VisualBasic.Information.VarType(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value) == global::Microsoft.VisualBasic.Information.VarType(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value)))
                                {
                                    if (((this.rng1CellValue.ToUpper() ?? "") != (this.rng2CellValue.ToUpper() ?? "")))
                                    {

                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                        label.ForeColor = this.CbFillFont.BackColor;

                                        if (((this.checkBoxFormatting.Checked) == (true)))
                                        {
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.BackColor = rgbColor;
                                        }
                                        else
                                        {
                                            label.BackColor = global::System.Drawing.Color.Transparent;

                                        }

                                        this.CP_Output_Range.Controls.Add(label);
                                    }
                                    else
                                    {
                                        var label = new global::System.Windows.Forms.Label();
                                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                        label.Height = (global::System.Int32)Math.Round(height);
                                        label.Width = (global::System.Int32)Math.Round(width);
                                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;


                                        if (((this.checkBoxFormatting.Checked) == (true)))
                                        {

                                            // background fill color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.BackColor = rgbColor;

                                            // font color
                                            lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                            rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                            label.ForeColor = rgbColor;
                                        }

                                        else
                                        {
                                            label.BackColor = global::System.Drawing.Color.Transparent;
                                            label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                        }

                                        this.CP_Output_Range.Controls.Add(label);

                                    }
                                }

                                else
                                {
                                    var label = new global::System.Windows.Forms.Label();
                                    label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                    label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                    label.Height = (global::System.Int32)Math.Round(height);
                                    label.Width = (global::System.Int32)Math.Round(width);
                                    label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;
                                    label.BackColor = global::System.Drawing.Color.Transparent;
                                    label.ForeColor = this.CbFillFont.BackColor;


                                    this.CP_Output_Range.Controls.Add(label);

                                }


                            }
                        }
                    }


                    // 1st range >> 2nd range >> radBtnDifferentValues checked >> case sensitive unchecked >> fill/font color not selected
                    else
                    {
                        for (global::System.Int32 i = 1, loopTo30 = displayRng.Rows.Count; i <= loopTo30; i++)
                        {
                            for (global::System.Int32 j = 1, loopTo31 = displayRng.Columns.Count; j <= loopTo31; j++)
                            {
                                this.rng1CellValue = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                this.rng2CellValue = Conversions.ToString(displayRng2.Cells[(global::System.Object)i, (global::System.Object)j].value);


                                var label = new global::System.Windows.Forms.Label();
                                label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round((global::System.Double)((((j) - (1)))) * (width)), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                label.Height = (global::System.Int32)Math.Round(height);
                                label.Width = (global::System.Int32)Math.Round(width);
                                label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;


                                if (((this.checkBoxFormatting.Checked) == (true)))
                                {

                                    // background fill color
                                    lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color);
                                    rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                    label.BackColor = rgbColor;

                                    // font color
                                    lblColor = Conversions.ToLong(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color);
                                    rgbColor = global::System.Drawing.Color.FromArgb((global::System.Int32)((lblColor) % (256L)), (global::System.Int32)(((((lblColor) / (256L)))) % (256L)), (global::System.Int32)(((((lblColor) / (65536L)))) % (256L)));
                                    label.ForeColor = rgbColor;
                                }

                                else
                                {
                                    label.BackColor = global::System.Drawing.Color.Transparent;
                                    label.ForeColor = (global::System.Drawing.Color)(default(global::System.Drawing.Color));

                                }


                                this.CP_Output_Range.Controls.Add(label);
                                // End If



                            }
                        }




                    }


                }



                this.CP_Output_Range.AutoScroll = true;
            }


            catch (global::System.Exception ex)
            {

            }

        }

        private void CBFillBackground_Click(global::System.Object sender, global::System.EventArgs e)
        {
            this.Display();
            if (((this.checkBoxFillBack.Checked) == (true)))
            {

                this.colorPick = this.CD_Fill_Background.ShowDialog();

                if ((this.colorPick == global::System.Windows.Forms.DialogResult.OK))
                {
                    this.CBFillBackground.BackColor = this.CD_Fill_Background.Color;
                    this.Display();

                }
                this.CP_Input_Range1.Focus();
            }
            else
            {
                this.CP_Input_Range1.Focus();


            }
        }


        private void CbFillFont_Click(global::System.Object sender, global::System.EventArgs e)
        {

            this.Display();
            if (((this.checkBoxFillFont.Checked) == (true)))
            {

                this.colorPick = this.CD_Fill_Font.ShowDialog();


                if ((this.colorPick == global::System.Windows.Forms.DialogResult.OK))
                {
                    this.CbFillFont.BackColor = this.CD_Fill_Font.Color;
                    this.Display();

                }
                this.CP_Input_Range1.Focus();
            }
            else
            {
                this.CP_Input_Range1.Focus();


            }

        }

        private void CBFillBackground_SelectedIndexChanged(global::System.Object sender, global::System.EventArgs e)
        {
            this.Display();

        }

        private void radBtnSameValues_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {
            this.Display();
        }

        private void radBtnDifferentValues_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {
            this.Display();
        }

        private void CustomPanel1_Paint(global::System.Object sender, global::System.Windows.Forms.PaintEventArgs e)
        {

        }

        private void checkBoxCase_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {
            this.Display();
        }

        private void CBFillBackground_BackColorChanged(global::System.Object sender, global::System.EventArgs e)
        {

            if (((this.CBFillBackground.BackColor.Name == "LightSteelBlue") & ((this.GB_Display_Result.BackColor) != (this.CBFillBackground.BackColor))))
            {

                return;

            }


            this.Display();



        }

        private void CbFillFont_BackColorChanged(global::System.Object sender, global::System.EventArgs e)
        {

            if (((this.CbFillFont.BackColor.Name == "MidnightBlue") & ((this.GB_Display_Result.BackColor) != (this.CBFillBackground.BackColor))))
            {

                return;

            }


            this.Display();



        }

        private void checkBoxFormatting_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {

            this.Display();


        }

        private void checkBoxFillBack_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {
            this.Display();

        }

        private void checkBoxFillFont_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {

            this.Display();

        }

        private void txtSourceRange1_Click(global::System.Object sender, global::System.EventArgs e)
        {
            // txtSourceRange1.SelectionStart = txtSourceRange1.TextLength
            // txtSourceRange1.ScrollToCaret()
        }

        private void txtSourceRange2_Click(global::System.Object sender, global::System.EventArgs e)
        {
            // txtSourceRange2.SelectionStart = txtSourceRange2.TextLength
            // txtSourceRange2.ScrollToCaret()
        }

        private void Form15CompareCells_Closing(global::System.Object sender, global::System.ComponentModel.CancelEventArgs e)
        {
            global::VSTO_Addins.GlobalModule.form_flag = false;
        }



        private void Form15CompareCells_Shown(global::System.Object sender, global::System.EventArgs e)
        {
            this.Focus();
            this.BringToFront();
            this.Activate();
            this.BeginInvoke(new global::System.Action(() =>
                {
                    this.txtSourceRange1.Text = this.firstInputRng.get_Address();
                    global::VSTO_Addins.Form15CompareCells.SetWindowPos(this.Handle, new global::System.IntPtr(global::VSTO_Addins.Form15CompareCells.HWND_TOPMOST), 0, 0, 0, 0, ((global::VSTO_Addins.Form15CompareCells.SWP_NOACTIVATE) | (global::VSTO_Addins.Form15CompareCells.SWP_NOMOVE)) | (global::VSTO_Addins.Form15CompareCells.SWP_NOSIZE));
                }));
        }

        private void Form15CompareCells_Disposed(global::System.Object sender, global::System.EventArgs e)
        {
            global::VSTO_Addins.GlobalModule.form_flag = false;
        }
    }
}