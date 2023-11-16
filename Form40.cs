using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using Point = System.Drawing.Point;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace VSTO_Addins
{

    public partial class Form40
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
        private Excel.Workbook workBook;
        public static Excel.Worksheet workSheet;

        private Range src_rng;
        public Range des_rng;
        private Range selectedRange;

        public Range validationRange;
        private List<string> allItems = new List<string>();

        private bool processingEvent = false;

        public Form40()
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



        private void Form40_Load(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            var cell = worksheet.get_Range(GlobalModule.TargetVar3); // In TargetVar, there is address about Target cell
            string validationFormula = cell.Validation.Formula1;
            var items = new List<string>();
            // MsgBox(validationFormula)
            // Dim items As New List(Of String)()

            if (!validationFormula.Contains(",") && !validationFormula.Contains("!"))
            {
                // It's a range on the same sheet
                var range = worksheet.get_Range(validationFormula);

                foreach (Range cellInRange in range.Cells)
                {
                    if (!string.IsNullOrEmpty(cellInRange.get_Value()?.ToString()))
                    {
                        items.Add(cellInRange.get_Value().ToString());
                        allItems.Add(cellInRange.get_Value().ToString()); // Add to the master list as well
                    }
                }
            }
            else if (validationFormula.Contains(","))
            {
                // Direct values separated by commas
                items.AddRange(validationFormula.Split(new char[] { ',' }));
                allItems.AddRange(validationFormula.Split(new char[] { ',' }));
            }

            ListBox1.Items.Clear();
            ListBox1.Items.AddRange(items.ToArray());

            BringToFront();


        }

        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            if (ListBox1.SelectedItem is not null)
            {
                // Set the value in B1 cell to the selected item
                string selectedItem = ListBox1.SelectedItem.ToString();
                worksheet.get_Range(GlobalModule.TargetVar3).set_Value(value: selectedItem);
            }
        }

        private void PictureBox3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            string searchTerm = txtSearch.Text.ToLower();

            // Filter items based on the search term
            var filteredItems = allItems.Where(item => item.ToLower().Contains(searchTerm)).ToList();

            // Update the ListBox
            ListBox1.Items.Clear();
            ListBox1.Items.AddRange(filteredItems.ToArray());


        }

        // For position of the form
        private void Form40_Activated(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;


            var excelWindow = excelApp.ActiveWindow;

            var cell = worksheet.get_Range(GlobalModule.TargetVar3).get_Offset(1, 1);
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
            Location = new Point(x, y) + (Size)new Point(2, 2);
            // MsgBox(Me.Location.ToString)

        }


        private void form_enter(object sender, KeyEventArgs e)
        {

            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Close();

                }
            }

            catch (Exception ex)
            {

            }

        }


        private void listbox_enter(object sender, KeyEventArgs e)
        {

            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Close();

                }
            }

            catch (Exception ex)
            {

            }

        }

        private void Form40_Shown(object sender, EventArgs e)
        {
            // Me.BringToFront()
            Focus();
        }

        private void Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void PictureBox4_Click(object sender, EventArgs e)
        {

        }
    }
}