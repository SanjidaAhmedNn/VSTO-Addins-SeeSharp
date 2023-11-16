using System;
using System.Collections.Generic;
using System.Data;
using DataTable = System.Data.DataTable;
using System.Drawing;
using Point = System.Drawing.Point;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace VSTO_Addins
{

    public partial class Form38
    {
        private DataTable dt;
        // Public dv As DataView
        // dim dv As New DataView(dt)

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
        private Excel.Workbook workBook;
        public static Excel.Worksheet workSheet;

        private Range src_rng;
        public Range des_rng;
        private Range selectedRange;

        public Range validationRange;
        public Form37_MSDropDownCheckBox form = null;

        private bool processingEvent = false;

        public Form38()
        {
            InitializeComponent();
        }
        // Public focuschange As Boolean

        [DllImport("user32")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private const uint SWP_NOMOVE = 0x2U;
        private const uint SWP_NOSIZE = 0x1U;
        private const uint SWP_NOACTIVATE = 0x10U;
        private const int HWND_TOPMOST = -1;

        // Public Target As Excel.Range


        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hWnd, ref RECT lpRect);


        private void Form38_Load(object sender, EventArgs e)
        {
            // Separator = ","
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;


            // Add Increment Button
            var incrementColumn = new DataGridViewCheckBoxColumn();


            incrementColumn.Name = "Increment";
            incrementColumn.HeaderText = "Check";

            incrementColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            incrementColumn.Width = 45; // Set the width you want here
                                        // incrementColumn.DefaultCellStyle.BackColor = Color.White
            incrementColumn.FlatStyle = FlatStyle.Popup;

            incrementColumn.DefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 12f);
            incrementColumn.ReadOnly = false;

            // .Columns(1).ReadOnly = True
            DataGridView1.Columns.Add(incrementColumn);


            // Dim formula1 As String
            // Dim dropdownItems() As String
            Range sourceRange;

            // Get the cell with the drop-down list
            var cell = worksheet.get_Range(GlobalModule.TargetVar2);


            // Populate DataGridView
            dt = new DataTable();
            // dt.Columns.Add("Value", GetType(String))
            var connectedstring = new List<string>();
            string connectedstringst = "";



            string validationList = "";
            string formula = cell.Validation.Formula1;
            if (formula.Contains(","))
            {
                // Data validation type: Excel Range
                validationList = formula;

                dt.Columns.Add("Value", typeof(string));
                dt.Rows.Add("Select all");
                string[] items = validationList.Split(',');
                foreach (string item in items)
                    dt.Rows.Add(item);
            }
            else
            {
                // Data validation type: Excel Range

                sourceRange = worksheet.get_Range(formula);

                dt.Columns.Add("Value", typeof(string));

                dt.Rows.Add("Select all");
                // Sample Data
                foreach (Range itemCell in sourceRange)
                    // dt.Rows.Add(20)
                    // dt.Rows.Add(30)
                    dt.Rows.Add(itemCell.get_Value());

            }


            DataGridView1.DataSource = dt;


            if (cell.get_Value() is not null)
            {
                connectedstringst = cell.get_Value().ToString();
            }


            // MsgBox(connectedstringst)

            // Parse the values
            // Dim values As String() = worksheet.Range(TargetVar).Split(","c)

            // DataGridView1.DataSource = dt
            int i = 0;

            if (connectedstringst is not null | !string.IsNullOrEmpty(connectedstringst))
            {

                foreach (DataGridViewRow r in DataGridView1.Rows)
                {

                    string cellValue = r.Cells[1].Value.ToString();

                    if (connectedstringst.Contains(cellValue.ToString()))
                    {
                        DataGridView1.Rows[i].Cells["Increment"].Value = true;
                    }
                    else
                    {
                        DataGridView1.Rows[i].Cells["Increment"].Value = false;
                    }
                    i = i + 1;
                }
            }

            // MsgBox(connectedstringst)

            // For Each r As DataGridViewRow In DataGridView1.Rows
            // If Not r.IsNewRow Then ' Avoid the last empty row
            // 'MsgBox(1)

            // Dim cellValue As Object = r.Cells(1).Value
            // If cellValue IsNot Nothing AndAlso connectedstringst.Contains(cellValue.ToString()) Then
            // r.Cells("Increment").Value = True
            // Else
            // r.Cells("Increment").Value = False
            // End If
            // End If
            // Next



            DataGridView1.Columns[1].Width = 190;
            string targetCellValue = Conversions.ToString(worksheet.get_Range(GlobalModule.TargetVar2).get_Value()); // Assuming B1 cell contains the values
                                                                                                                     // Me.BringToFront()
            TopMost = true;
            // Me.Show()
            TopMost = false;

        }


        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            // Ensure it's not the header row
            if (e.RowIndex < 0)
                return;

            var cell = DataGridView1.Rows[e.RowIndex].Cells["Value"];


            // ' Check for Increment button click
            // If e.ColumnIndex = DataGridView1.Columns("Increment").Index Then
            // cell.Value = Convert.ToInt32(cell.Value) + 1
            // End If

            // ' Check for Decrement button click
            // If e.ColumnIndex = DataGridView1.Columns("Decrement").Index Then
            // cell.Value = Convert.ToInt32(cell.Value) - 1
            // End If
            if (Conversions.ToBoolean(Operators.AndObject(Operators.AndObject(e.ColumnIndex == DataGridView1.Columns["Increment"].Index, Operators.ConditionalCompareObjectEqual(DataGridView1.Rows[e.RowIndex].Cells["Increment"].Value, false, false)), Operators.ConditionalCompareObjectEqual(DataGridView1.Rows[e.RowIndex].Cells[1].Value, "Select all", false))))
            {
                // DataGridView1.Rows(e.RowIndex).Cells("Increment").Value = True



                foreach (DataGridViewRow r in DataGridView1.Rows)



                    // DataGridView1.Rows(i).Cells("Increment").Value = True
                    // DataGridView1.Rows(j).Cells("Increment") = True
                    // MsgBox(1)


                    r.Cells[0].Value = true;
            }




            else if (Conversions.ToBoolean(Operators.AndObject(Operators.AndObject(e.ColumnIndex == DataGridView1.Columns["Increment"].Index, Operators.ConditionalCompareObjectEqual(DataGridView1.Rows[e.RowIndex].Cells["Increment"].Value, true, false)), Operators.ConditionalCompareObjectEqual(DataGridView1.Rows[e.RowIndex].Cells[1].Value, "Select all", false))))
            {
                // DataGridView1.Rows(e.RowIndex).Cells("Increment").Value = False



                foreach (DataGridViewRow r in DataGridView1.Rows)

                    // Dim cellValue As String = r.Cells(1).Value.ToString()

                    // DataGridView1.Rows(i).Cells("Increment").Value = True
                    // DataGridView1.Rows(j).Cells("Increment") = True
                    // MsgBox(1)


                    r.Cells[0].Value = false;
            }
            // End If

            else if (Conversions.ToBoolean(Operators.AndObject(Operators.AndObject(e.ColumnIndex == DataGridView1.Columns["Increment"].Index, Operators.ConditionalCompareObjectEqual(DataGridView1.Rows[e.RowIndex].Cells["Increment"].Value, false, false)), Operators.ConditionalCompareObjectNotEqual(DataGridView1.Rows[e.RowIndex].Cells[1].Value, "Select all", false))))
            {


                DataGridView1.Rows[e.RowIndex].Cells["Increment"].Value = true; // or False
                if (worksheet.get_Range(GlobalModule.TargetVar2).get_Value() is null)
                {
                    worksheet.get_Range(GlobalModule.TargetVar2).set_Value(value: cell.Value);
                }
                else
                {
                    string itemToRemove = Conversions.ToString(cell.Value);


                }
            }

            else if (Conversions.ToBoolean(Operators.AndObject(Operators.AndObject(e.ColumnIndex == DataGridView1.Columns["Increment"].Index, Operators.ConditionalCompareObjectEqual(DataGridView1.Rows[e.RowIndex].Cells["Increment"].Value, true, false)), Operators.ConditionalCompareObjectNotEqual(DataGridView1.Rows[e.RowIndex].Cells[1].Value, "Select all", false))))
            {
                DataGridView1.Rows[e.RowIndex].Cells["Increment"].Value = false;
                // Me.Refresh()
                string itemToRemove = Conversions.ToString(cell.Value);

                if (worksheet.get_Range(GlobalModule.TargetVar2).get_Value() is not null && worksheet.get_Range(GlobalModule.TargetVar2).get_Value().ToString().Contains(itemToRemove))
                {



                    var items = worksheet.get_Range(GlobalModule.TargetVar2).get_Value().ToString().Split(new string[] { GlobalModule.Separator2 }, StringSplitOptions.None).ToList();

                    // Find the index of the first occurrence of the item to remove
                    int indexToRemove = items.FindIndex(x => (x.Trim() ?? "") == (itemToRemove ?? ""));

                    if (indexToRemove >= 0) // If found
                    {
                        items.RemoveAt(indexToRemove); // Remove only the first occurrence
                        worksheet.get_Range(GlobalModule.TargetVar2).set_Value(value: string.Join(GlobalModule.Separator2, items));
                    }
                }

            }




        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            // Dim searchTerm As String = txtSearch.Text.Trim()
            // If String.IsNullOrEmpty(searchTerm) Then
            // DataGridView1.DataSource = dt
            // Else
            // Dim dv As New DataView(dt)

            // If IsNumeric(searchTerm) Then
            // dv.RowFilter = String.Format("Value = {0}", Convert.ToInt32(searchTerm))
            // DataGridView1.DataSource = dv
            // Else
            // 'MessageBox.Show("Please enter a valid number.")
            // End If
            // End If
        }

        // dv.RowFilter = String.Format("Convert(Value, 'System.String') LIKE '{0}%'", searchTerm)

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {

            string searchTerm = txtSearch.Text.Trim();

            if (string.IsNullOrEmpty(searchTerm))
            {
                DataGridView1.DataSource = dt;
            }
            else
            {
                var dv = new DataView(dt);

                if (Information.IsNumeric(searchTerm))
                {
                    dv.RowFilter = string.Format("Convert(Value, 'System.String') LIKE '{0}%'", searchTerm);
                    DataGridView1.DataSource = dv;
                }
                else
                {
                    DataGridView1.DataSource = dt;
                }
            }


            // Dim searchTerm As String = txtSearch.Text.Trim()
            // 'Dim dv As New DataView(dt)

            // If String.IsNullOrEmpty(searchTerm) Then

            // dv.RowFilter = "" ' Clear the filter
            // Else
            // If IsNumeric(searchTerm) Then
            // dv.RowFilter = String.Format("Convert(Value, 'System.String') LIKE '{0}%'", searchTerm)
            // Else
            // dv.RowFilter = "" ' Clear the filter if search term is non-numeric
            // End If
            // End If
        }

        private void PictureBox3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Form38_Activated(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;


            var excelWindow = excelApp.ActiveWindow;

            var cell = worksheet.get_Range(GlobalModule.TargetVar2).get_Offset(1, 1);
            var zoomFactor = Operators.DivideObject(excelWindow.Zoom, 100);
            // var ws = cell.Worksheet;

            var ap = excelWindow.ActivePane; // might be split panes
            int origScrollCol = ap.ScrollColumn;
            int origScrollRow = ap.ScrollRow;
            excelApp.ScreenUpdating = false;
            // when FreezePanes == true, ap.ScrollColumn/Row will only reset
            // as much as the location of the frozen splitter
            ap.ScrollColumn = 1;
            ap.ScrollRow = 1;

            // PointsToScreenPixels returns different values if the scroll Is Not currently 1
            // Temporarily set the scroll back to 1 so that PointsToScreenPixels returns a
            // value we know how to handle.
            // (x,y) are screen coordinates for the top left corner of the top left cell
            int x = ap.PointsToScreenPixelsX(0); // e.g. window.x + row header width
            int y = ap.PointsToScreenPixelsY(0); // e.g. window.y + ribbon height + column headers height

            float dpiX = 0f;
            float dpiY = 0f;

            using (var g = Graphics.FromHwnd(IntPtr.Zero))
            {
                dpiX = g.DpiX;
                dpiY = g.DpiY;
            }

            int deltaRow = 0;
            int deltaCol = 0;
            int fromCol = origScrollCol;
            int fromRow = origScrollRow;
            if (excelWindow.FreezePanes)
            {
                fromCol = 1;
                fromRow = 1;
                deltaCol = origScrollCol - ap.ScrollColumn; // // Note: ap.ScrollColumn/ Row <> 1
                deltaRow = origScrollRow - ap.ScrollRow;  // // see comment: when FreezePanes == true ...
            }

            // // Note Each column width / row height has to be calculated individually.
            // // Before, tried to use this approach:
            // // var r2 = (Microsoft.Office.Interop.Excel.Range) cell.Worksheet.Cells[origScrollRow, origScrollCol];
            // // double dw = cell.Left - r2.Left;
            // // double dh = cell.Top - r2.Top;
            // // However, that only works when the zoom factor Is a whole number.
            // // A fractional zoom (e.g. 1.27) causes each individual row Or column to round to the closest whole number,
            // // which means having to loop through.

            Range col;
            double ww;
            double newW;
            int i;
            var loopTo = cell.Column - 1;
            for (i = fromCol; i <= loopTo; i++)
            {
                // skip the columns between the frozen split and the first visible column
                if (i >= ap.ScrollColumn && i < ap.ScrollColumn + deltaCol)
                {
                    continue;
                }

                col = (Range)worksheet.Cells[cell.Row, i];
                ww = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(col.Width, dpiX), 72));
                newW = Conversions.ToDouble(Operators.MultiplyObject(zoomFactor, ww));
                x += (int)Math.Round(Math.Round(newW));
            }


            Range row;
            double hh;
            double newH;

            var loopTo1 = cell.Row - 1;
            for (i = fromRow; i <= loopTo1; i++)
            {
                // skip the rows between the frozen split and the first visible row
                if (i >= ap.ScrollRow && i < ap.ScrollRow + deltaRow)
                {
                    continue;
                }

                row = (Range)worksheet.Cells[i, cell.Column];
                hh = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(row.Height, dpiY), 72));
                newH = Conversions.ToDouble(Operators.MultiplyObject(zoomFactor, hh));
                y += (int)Math.Round(Math.Round(newH));
            }

            ap.ScrollColumn = origScrollCol;
            ap.ScrollRow = origScrollRow;
            excelApp.ScreenUpdating = true;


            Location = new Point(x, y) + (Size)new Point(2, 2);
            // MsgBox(Me.Location.ToString)

        }

        private void DataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {


            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            // Ensure it's not the header row
            if (e.RowIndex < 0)
                return;

            // Dim cell As DataGridViewCell = DataGridView1.Rows(e.RowIndex).Cells("Value")


            bool isChecked = Conversions.ToBoolean(DataGridView1.Rows[e.RowIndex].Cells[0].Value);
            string itemValue = DataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            // MsgBox(1)
            // If isChecked Then
            // If itemValue = "Select all" And isChecked = True Then


            // Dim i As Integer = 0
            // ' Dim j As Integer = 1

            // For Each r As DataGridViewRow In DataGridView1.Rows

            // Dim cellValue As String = r.Cells(1).Value.ToString()

            // 'DataGridView1.Rows(i).Cells("Increment").Value = True
            // r.Cells(0).Value = True
            // 'DataGridView1.Rows(j).Cells("Increment") = True
            // 'MsgBox(1)

            // i = i + 1
            // Next




            // ' DataGridView1.Columns("Increment").cells(0).value = True
            // 'DataGridView1.Rows(e.RowIndex).Cells("Increment").Value = True

            // Else

            // Dim i As Integer = 0
            // ' Dim j As Integer = 1

            // For Each r As DataGridViewRow In DataGridView1.Rows

            // 'Dim cellValue As String = r.Cells(1).Value.ToString()

            // r.Cells(0).Value = False
            // 'MsgBox(2)
            // 'DataGridView1.Rows(j).Cells("Increment") = True

            // i = i + 1
            // Next
            // Me.Refresh()
            // End If

            if (isChecked == true & itemValue != "Select all")
            {
                Refresh();
                // Place the item in B1 cell
                // worksheet.Range("B1").Value = DataGridView1.Rows(e.RowIndex).Cells("YourItemColumnName").Value
                if (worksheet.get_Range(GlobalModule.TargetVar2).get_Value() is null)
                {
                    worksheet.get_Range(GlobalModule.TargetVar2).set_Value(value: itemValue);
                }
                else
                {
                    string values;
                    try
                    {

                        // values = worksheet.Range(TargetVar2).Value.Split(","c)
                        // values = worksheet.Range(TargetVar2).Value.ToString.Split(New String Separator2, StringSplitOptions.None)
                        // values = worksheet.Range(TargetVar2).Value.ToString().Split(String() Separator2, StringSplitOptions.None)

                        // Split the string using the separator
                        string[] result = worksheet.get_Range(GlobalModule.TargetVar2).get_Value().ToString().Split(new string[] { GlobalModule.Separator2 }, StringSplitOptions.None);

                        // Join the split results back into a single string, using a space as the new separator (or any other separator of your choice)
                        values = string.Join(" ", result);
                    }
                    catch (Exception ex)
                    {
                        values = Conversions.ToString(worksheet.get_Range(GlobalModule.TargetVar2).get_Value());
                    }
                    if (values.Contains(itemValue) == false & GlobalModule.Horizontal2 == true)
                    {

                        // If Horizontal = True Then
                        worksheet.get_Range(GlobalModule.TargetVar2).set_Value(value: Operators.ConcatenateObject(Operators.ConcatenateObject(worksheet.get_Range(GlobalModule.TargetVar2).get_Value(), GlobalModule.Separator2), itemValue));
                    }
                    else if (values.Contains(itemValue) == false & GlobalModule.Horizontal2 == false)
                    {
                        worksheet.get_Range(GlobalModule.TargetVar2).set_Value(value: Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(worksheet.get_Range(GlobalModule.TargetVar2).get_Value(), GlobalModule.Separator2), Microsoft.VisualBasic.Constants.vbNewLine), itemValue));
                    }
                }
            }
            // ElseIf e.ColumnIndex = DataGridView1.Columns("Decrement").Index AndAlso e.RowIndex >= 0 Then
            // Me.Refresh()
            // Dim itemToRemove As String = cell.Value

            // If worksheet.Range(TargetVar).Value IsNot Nothing AndAlso worksheet.Range(TargetVar).Value.ToString().Contains(itemToRemove) Then
            // Dim items As List(Of String) = worksheet.Range(TargetVar).Value.ToString().Split(Separator).ToList()
            // items.RemoveAll(Function(x) x.Trim() = itemToRemove)
            // worksheet.Range(TargetVar).Value = String.Join(Separator, items)
            // End If

            else if (isChecked == false)
            {
                // Me.Refresh()
                string itemToRemove = itemValue;

                if (worksheet.get_Range(GlobalModule.TargetVar2).get_Value() is not null && worksheet.get_Range(GlobalModule.TargetVar2).get_Value().ToString().Contains(itemToRemove))
                {
                    var items = worksheet.get_Range(GlobalModule.TargetVar2).get_Value().ToString().Split(new string[] { GlobalModule.Separator2 }, StringSplitOptions.None).ToList();

                    // Find the index of the first occurrence of the item to remove
                    int indexToRemove = items.FindIndex(x => (x.Trim() ?? "") == (itemToRemove ?? ""));

                    if (indexToRemove >= 0) // If found
                    {
                        items.RemoveAt(indexToRemove); // Remove only the first occurrence
                        worksheet.get_Range(GlobalModule.TargetVar2).set_Value(value: string.Join(GlobalModule.Separator2, items));
                    }
                }

            }


            // If e.ColumnIndex = 0 Then
            // Dim isChecked As Boolean = DataGridView1.Rows(e.RowIndex).Cells(0).Value
            // Dim itemValue As String = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString()
            // MsgBox(1)
            // If isChecked Then
            // AddToExcelDropdownList(itemValue)
            // MsgBox(2)
            // Else
            // RemoveFromExcelDropdownList(itemValue)
            // MsgBox(3)
            // End If
            // End If
        }

        private void AddToExcelDropdownList(string value)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            Range cell;
            // Dim existingValues As String

            // workbook = excelApp.Workbooks.Open("YOUR_EXCEL_PATH_HERE.xlsx")
            // worksheet = workbook.Worksheets(1)
            Interaction.MsgBox(4);
            cell = (Range)excelApp.Cells[1, 2]; // Change this to your dropdown cell location

            // Assuming the cell has a dropdown list validation
            // existingValues = cell.Validation.Formula1

            // If Not existingValues.Contains(value) Then
            worksheet.get_Range(GlobalModule.TargetVar2).set_Value(value: Operators.ConcatenateObject(Operators.ConcatenateObject(worksheet.get_Range(GlobalModule.TargetVar2).get_Value(), ","), value));

            // MsgBox(5)

        }

        private void RemoveFromExcelDropdownList(string value)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            Range cell;

            cell = (Range)excelApp.Cells[1, 2]; // Change this to your dropdown cell location

            string itemToRemove = Conversions.ToString(cell.get_Value());

            if (worksheet.get_Range(GlobalModule.TargetVar2).get_Value() is not null && worksheet.get_Range(GlobalModule.TargetVar2).get_Value().ToString().Contains(itemToRemove))
            {
                var items = worksheet.get_Range(GlobalModule.TargetVar2).get_Value().ToString().Split(new string[] { GlobalModule.Separator2 }, StringSplitOptions.None).ToList();

                // Find the index of the first occurrence of the item to remove
                int indexToRemove = items.FindIndex(x => (x.Trim() ?? "") == (itemToRemove ?? ""));

                if (indexToRemove >= 0) // If found
                {
                    items.RemoveAt(indexToRemove); // Remove only the first occurrence
                    worksheet.get_Range(GlobalModule.TargetVar2).set_Value(value: string.Join(GlobalModule.Separator2, items));
                }
            }

        }

        private void Form38_Shown(object sender, EventArgs e)
        {
            // Me.BringToFront()
            Focus();
        }

        private void PictureBox4_Click(object sender, EventArgs e)
        {
            GlobalModule.settingflag2 = true;
            Hide();
            form = new Form37_MSDropDownCheckBox();
            form.Show();
            form.CustomGroupBox6.Enabled = false;
            if (form is null | form.IsDisposed == true)
            {
                Show();
            }
        }

        private void Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}