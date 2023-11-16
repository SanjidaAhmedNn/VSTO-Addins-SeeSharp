using System;
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

    public partial class Form36
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
        private bool power;

        private bool processingEvent = false;
        public bool focuschange;
        private Form35Multi_SelectionbasedDropdown Form = null;

        public Form36()
        {
            InitializeComponent();
        }

        [DllImport("user32")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private const uint SWP_NOMOVE = 0x2U;
        private const uint SWP_NOSIZE = 0x1U;
        private const uint SWP_NOACTIVATE = 0x10U;
        private const int HWND_TOPMOST = -1;

        public Range Target;


        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hWnd, ref RECT lpRect);


        private void Form36_Load(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;


            // Enable & Disable Search Option
            if (GlobalModule.Search1 == true)
            {
                txtSearch.Enabled = true;
                PB_Search.Enabled = true;
            }
            else
            {
                txtSearch.Enabled = false;
                PB_Search.Enabled = false;
            }


            // Add Increment Button
            var incrementColumn = new DataGridViewButtonColumn();
            incrementColumn.Name = "Increment";
            incrementColumn.HeaderText = "Add";
            incrementColumn.Text = "+";
            incrementColumn.UseColumnTextForButtonValue = true;
            incrementColumn.Width = 28; // Set the width you want here
            incrementColumn.DefaultCellStyle.BackColor = Color.White;
            incrementColumn.FlatStyle = FlatStyle.Popup;
            // MsgBox(incrementColumn.DefaultCellStyle.BackColor.ToString)
            incrementColumn.DefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 12f);
            DataGridView1.Columns.Add(incrementColumn);

            // Add Decrement Button
            var decrementColumn = new DataGridViewButtonColumn();
            decrementColumn.Name = "Decrement";
            decrementColumn.HeaderText = "Sub";
            decrementColumn.Text = "-";
            decrementColumn.UseColumnTextForButtonValue = true;
            decrementColumn.Width = 28; // Set the width you want here
            decrementColumn.DefaultCellStyle.BackColor = Color.White;
            decrementColumn.FlatStyle = FlatStyle.Popup;
            decrementColumn.DefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 12f);
            DataGridView1.Columns.Add(decrementColumn);



            // Dim formula1 As String
            // Dim dropdownItems() As String
            Range sourceRange = null;

            // Get the cell with the drop-down list
            var cell = worksheet.get_Range(GlobalModule.TargetVar1);


            // Extract the formula (assuming it's a list)
            // formula1 = cell.Validation.Formula1
            // dropdownItems = formula1.Split(","c)


            dt = new DataTable();
            string validationList = "";
            string formula = cell.Validation.Formula1;

            if (formula.Contains(","))
            {
                // Data validation type: Excel Range
                validationList = formula;

                dt.Columns.Add("Value", typeof(string));

                string[] items = validationList.Split(',');
                // dt.Rows(1).Add("Select All")
                foreach (string item in items)
                    // dt.Rows.Add(20)
                    dt.Rows.Add(item);
            }
            // dt.Rows.Add(20)
            else
            {
                // Data validation type: Excel Range

                sourceRange = worksheet.get_Range(formula);

                dt.Columns.Add("Value", typeof(string));

                // dt.Rows.Add("Select All")
                // Sample Data
                foreach (Range itemCell in sourceRange)
                    // dt.Rows.Add(20)
                    // dt.Rows.Add(30)
                    dt.Rows.Add(itemCell.get_Value());
                // dt.Rows.Add(20)

            }









            // ' Extract the formula (assuming it's a range reference)
            // formula1 = cell.Validation.Formula1
            // ' Assuming the range is in the same sheet
            // sourceRange = worksheet.Range(formula1)

            // Populate DataGridView
            // For Each item As String In dropdownItems
            // DataGridView1.Rows.Add(item)
            // Next


            // DataGridView1.Rows.Add("Select All")
            DataGridView1.DataSource = dt;
            DataGridView1.Columns[2].Width = 110;


            var labelColumn = new DataGridViewTextBoxColumn();
            labelColumn.Name = "OccurrenceCount";
            labelColumn.HeaderText = "Occurrences";
            DataGridView1.Columns.Add(labelColumn);
            labelColumn.Width = 72;
            labelColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            // Dim excelApp As New Microsoft.Office.Interop.Excel.Application
            // Dim ws As Microsoft.Office.Interop.Excel.Worksheet = excelApp.ActiveSheet
            // Dim targetRange As Excel.Range = ws.Range(rangeAddress)
            string targetCellValue = Conversions.ToString(worksheet.get_Range(GlobalModule.TargetVar1).get_Value()); // Assuming B1 cell contains the values

            foreach (DataGridViewRow r in DataGridView1.Rows)
            {
                string itemValue = r.Cells[2].Value.ToString();
                // MsgBox(itemValue)
                int count = CountOccurrencesInExcelCell(itemValue, targetCellValue);

                r.Cells["OccurrenceCount"].Value = count;
            }

            // DataGridView1.DataSource = dv

            BringToFront();
            Focus();


        }

        private int CountOccurrencesInExcelCell(string item, string cellValue)
        {
            // If String.IsNullOrEmpty(cellValue) Then Return 0

            // ' Split the cellValue using a comma and remove any extra white spaces
            // Dim items As String() = cellValue.Split(New Char() {","c}).Select(Function(s) s.Trim()).ToArray()

            // ' Count occurrences of the item in the split array
            // Return items.Count(Function(i) i = item)
            int occurrences = 0;

            if (cellValue is not null)
            {
                // MsgBox(1)
                // Dim cellValue As String = targetCell.Value.ToString()
                string[] items = cellValue.Split(new char[] { Conversions.ToChar(GlobalModule.Separator1) }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string i in items)
                {
                    if (i.Trim().Equals(item.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        occurrences += 1;
                    }
                }
            }
            // Me.Refresh()
            return occurrences;
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


            if (e.ColumnIndex == DataGridView1.Columns["Increment"].Index)
            {
                Refresh();
                // Place the item in B1 cell
                // worksheet.Range("B1").Value = DataGridView1.Rows(e.RowIndex).Cells("YourItemColumnName").Value
                if (worksheet.get_Range(GlobalModule.TargetVar1).get_Value() is null)
                {
                    worksheet.get_Range(GlobalModule.TargetVar1).set_Value(value: cell.Value);
                }
                else if (GlobalModule.Horizontal1 == true)
                {
                    worksheet.get_Range(GlobalModule.TargetVar1).set_Value(value: Operators.ConcatenateObject(Operators.ConcatenateObject(worksheet.get_Range(GlobalModule.TargetVar1).get_Value(), GlobalModule.Separator1), cell.Value));
                }
                else
                {
                    worksheet.get_Range(GlobalModule.TargetVar1).set_Value(value: Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(worksheet.get_Range(GlobalModule.TargetVar1).get_Value(), GlobalModule.Separator1), Microsoft.VisualBasic.Constants.vbNewLine), cell.Value));
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

            else if (e.ColumnIndex == DataGridView1.Columns["Decrement"].Index && e.RowIndex >= 0)
            {
                // Me.Refresh()
                string itemToRemove = Conversions.ToString(cell.Value);

                if (worksheet.get_Range(GlobalModule.TargetVar1).get_Value() is not null && worksheet.get_Range(GlobalModule.TargetVar1).get_Value().ToString().Contains(itemToRemove))
                {
                    var items = worksheet.get_Range(GlobalModule.TargetVar1).get_Value().ToString().Split(new string[] { GlobalModule.Separator1 }, StringSplitOptions.None).ToList();

                    // Find the index of the first occurrence of the item to remove
                    int indexToRemove = items.FindIndex(x => (x.Trim() ?? "") == (itemToRemove ?? ""));

                    if (indexToRemove >= 0) // If found
                    {
                        items.RemoveAt(indexToRemove); // Remove only the first occurrence
                        worksheet.get_Range(GlobalModule.TargetVar1).set_Value(value: string.Join(GlobalModule.Separator1, items));
                    }
                }

            }


            // For Each r As DataGridViewRow In DataGridView1.Rows
            // If txtSearch Is Nothing Then
            // Dim itemValue As String = r.Cells(2).Value.ToString()
            // MsgBox(itemValue)
            // Dim count As Integer = CountOccurrencesInExcelCell(itemValue, worksheet.Range(TargetVar).Value)

            // r.Cells("OccurrenceCount").Value = count

            // End If
            // Next



            // If txtSearch.Text = "" Then
            foreach (DataGridViewRow r in DataGridView1.Rows)
            {
                // MsgBox(DataGridView1.Rows.Count)
                string itemValue = r.Cells[2].Value.ToString();
                if (power == true)
                {
                    itemValue = r.Cells[3].Value.ToString();
                }
                // MsgBox(itemValue)
                int count = CountOccurrencesInExcelCell(itemValue, Conversions.ToString(worksheet.get_Range(GlobalModule.TargetVar1).get_Value()));

                r.Cells["OccurrenceCount"].Value = count;
            }
            // Else
            // For Each r As DataGridViewRow In dt.Rows
            // Dim itemValue As String = r.Cells(2).Value.ToString()
            // 'MsgBox(itemValue)
            // Dim count As Integer = CountOccurrencesInExcelCell(itemValue, worksheet.Range(TargetVar).Value)

            // r.Cells("OccurrenceCount").Value = count
            // Next

            // End If

            Refresh();
        }

        // Private Sub btnSearch_Click(sender As Object, e As EventArgs)
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
        // End Sub

        // dv.RowFilter = String.Format("Convert(Value, 'System.String') LIKE '{0}%'", searchTerm)

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            if (string.IsNullOrEmpty(txtSearch.Text) | string.IsNullOrEmpty(txtSearch.Text))
            {
                power = false;
            }
            else
            {
                power = true;
            }


            string searchTerm = txtSearch.Text.Trim();
            var dv = new DataView(dt);
            if (string.IsNullOrEmpty(searchTerm))
            {
                DataGridView1.DataSource = dt;
            }
            // Dim dv As New DataView(dt)

            else if (Information.IsNumeric(searchTerm))
            {
                dv.RowFilter = string.Format("Convert(Value, 'System.String') LIKE '{0}%'", searchTerm);

                DataGridView1.DataSource = dv;
            }
            else
            {
                DataGridView1.DataSource = dt;
            }
            Refresh();
            // MsgBox(DataGridView1.Rows(1).Cells(2).Value.ToString)

            foreach (DataGridViewRow r in DataGridView1.Rows)
            {
                Refresh();
                // MsgBox(r.Cells(3).Value)
                string itemValue = r.Cells[3].Value.ToString();
                // MsgBox(itemValue)
                int count = CountOccurrencesInExcelCell(itemValue, Conversions.ToString(worksheet.get_Range(GlobalModule.TargetVar1).get_Value()));

                r.Cells["OccurrenceCount"].Value = count;
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
            Refresh();
        }

        private void PictureBox3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Form36_Activated(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;


            var excelWindow = excelApp.ActiveWindow;

            var cell = worksheet.get_Range(GlobalModule.TargetVar1).get_Offset(1, 1);
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

            // myFormInstance = New Form36()
            // yourFormInstance.Show()

            // Form f = New Form();
            // Me.Show()
            // Me.StartPosition = FormStartPosition.Manual
            BringToFront();
            Focus();
            Activate();
            Location = new Point(x, y) + (Size)new Point(2, 2);
            // MsgBox(Me.Location.ToString)


            TopMost = true; // Then it will bring the form to top
            TopMost = false;

        }


        private void PictureBox4_Click(object sender, EventArgs e)
        {
            GlobalModule.settingflag1 = true;
            Hide();
            Form = new Form35Multi_SelectionbasedDropdown();
            Form.Show();
            Form.CustomGroupBox6.Enabled = false;
            if (Form is null | Form.IsDisposed == true)
            {
                Show();
            }

        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Form36_Shown(object sender, EventArgs e)
        {
            // Me.BringToFront()
            Focus();
        }

        private void Panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}