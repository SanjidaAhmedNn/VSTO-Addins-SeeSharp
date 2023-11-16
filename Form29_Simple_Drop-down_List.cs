using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;

namespace VSTO_Addins
{



    public partial class Form29_Simple_Drop_down_List
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
                if (_excelApp != null)
                {
                    _excelApp.SheetSelectionChange -= excelApp_SheetSelectionChange;
                }

                _excelApp = value;
                if (_excelApp != null)
                {
                    _excelApp.SheetSelectionChange += excelApp_SheetSelectionChange;
                }
            }
        }
        private Excel.Workbook workBook;
        private Excel.Worksheet workSheet;
        private Excel.Worksheet workSheet2;
        private Excel.Range src_rng;
        public Excel.Range des_rng;
        private Excel.Range selectedRange;
        public bool focuschange = false;
        private string ax;

        public Form29_Simple_Drop_down_List()
        {
            InitializeComponent();
        }

        [DllImport("user32")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private const uint SWP_NOMOVE = 0x2U;
        private const uint SWP_NOSIZE = 0x1U;
        private const uint SWP_NOACTIVATE = 0x10U;
        private const int HWND_TOPMOST = -1;


        private int opened;
        private void Info_Click(object sender, EventArgs e)
        {

        }

        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Clear the list box
            List_Preview.Items.Clear();
            string selectedItem = ListBox1.SelectedItem.ToString();
            // Split the string into an array of strings
            string[] items = selectedItem.Split(',');

            List_Preview.Items.AddRange(items);
            Label7.Visible = true;
            Label7.Text = items.Count().ToString();

        }


        private void ComboBox1_MouseClick(object sender, MouseEventArgs e)
        {
            if (string.IsNullOrEmpty(ComboBox1.Text))
            {
            }
            // Do nothing
            else
            {
                // Clear the list box
                List_Preview.Items.Clear();
                string selectedItem = ComboBox1.Text;
                // Split the string into an array of strings
                string[] items = selectedItem.Split(',');

                List_Preview.Items.AddRange(items);
                Label7.Visible = true;
                Label7.Text = items.Count().ToString();
            }
        }

        private void ComboBox1_Enter(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(ComboBox1.Text))
            {
            }
            // Do nothing
            else
            {
                // Clear the list box
                List_Preview.Items.Clear();
                string selectedItem = ComboBox1.Text;
                // Split the string into an array of strings
                string[] items = selectedItem.Split(',');

                List_Preview.Items.AddRange(items);
                Label7.Visible = true;
                Label7.Text = items.Count().ToString();
            }
        }


        private void ComboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (string.IsNullOrEmpty(ComboBox1.Text))
            {
            }
            // Do nothing
            else
            {
                // Clear the list box
                List_Preview.Items.Clear();
                string selectedItem = ComboBox1.Text;
                // Split the string into an array of strings
                string[] items = selectedItem.Split(',');

                for (int i = 0, loopTo = items.Length - 1; i <= loopTo; i++)
                    items[i] = items[i].TrimStart();


                // ComboBox1.Items.AddRange(items)
                List_Preview.Items.AddRange(items);
                Label7.Visible = true;
                Label7.Text = items.Count().ToString();
            }
        }

        private void ComboBox1_KeyDown(object sender, KeyEventArgs e)
        {
            // Check if the key pressed was 'Enter'
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Btn_OK_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void ComboBox1_Leave(object sender, EventArgs e)
        {
            AddNewItem(ComboBox1.Text);
        }

        private void AddNewItem(string item)
        {
            // Check if the item is not already in the ComboBox
            if (!ComboBox1.Items.Contains(item))
            {
                ComboBox1.Items.Add(item);
            }
        }
        private void Selection()
        {
            if (string.IsNullOrEmpty(ComboBox1.Text))
            {
            }
            // Do nothing
            else
            {
                // Clear the list box
                List_Preview.Items.Clear();
                string selectedItem = ComboBox1.Text;
                // Split the string into an array of strings
                string[] items = selectedItem.Split(',');

                for (int i = 0, loopTo = items.Length - 1; i <= loopTo; i++)
                    items[i] = items[i].TrimStart();


                // ComboBox1.Items.AddRange(items)
                List_Preview.Items.AddRange(items);
                Label7.Visible = true;
                Label7.Text = items.Count().ToString();
            }
        }



        public void Btn_OK_Click(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            workBook = excelApp.ActiveWorkbook;
            workSheet = (Excel.Worksheet)workBook.ActiveSheet;
            try
            {

                if (string.IsNullOrEmpty(TB_dest_range.Text))
                {

                    if (RadioButton1.Checked == true & string.IsNullOrEmpty(TB_src_range.Text))
                    {
                        MessageBox.Show("Please Provide all inputs.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        TB_dest_range.Focus();
                        return;
                    }

                    else if (RadioButton2.Checked == true & ListBox1.SelectedIndex == -1)
                    {
                        MessageBox.Show("Please Provide all inputs.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        TB_dest_range.Focus();
                        return;
                    }

                    else if (RadioButton3.Checked == true & string.IsNullOrEmpty(ComboBox1.Text)) // No item selected
                    {
                        MessageBox.Show("Please Provide all inputs.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        TB_dest_range.Focus();
                        return;
                    }
                    else
                    {
                        MessageBox.Show("Select the Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        TB_dest_range.Focus();

                        return;
                    }
                }


                else if (IsValidExcelCellReference(TB_dest_range.Text) == false)
                {
                    MessageBox.Show("Please Enter valid Destination Range", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TB_dest_range.Focus();
                    return;
                }

                else if (RadioButton1.Checked == true & IsValidExcelCellReference(TB_src_range.Text) == false)
                {
                    MessageBox.Show("Please Enter valid Source Range", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TB_src_range.Focus();
                    return;
                }


                else if (RadioButton2.Checked == true & ListBox1.SelectedIndex == -1) // No item selected
                {
                    // Show message box to the user
                    MessageBox.Show("Please Select one Predefined Lists", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    ListBox1.Focus();
                    return;
                }

                else if (RadioButton3.Checked == true & string.IsNullOrEmpty(ComboBox1.Text)) // No item selected
                {
                    // Show message box to the user
                    MessageBox.Show("Please provide Source Range", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    ComboBox1.Focus();
                    return;
                }

                else if (RadioButton1.Checked == true)
                {

                    if (src_rng.Areas.Count > 1)
                    {
                        MessageBox.Show("Multiple selection is not possible in the Source Range field.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        TB_src_range.Focus();
                        return;
                    }

                    else if (string.IsNullOrEmpty(TB_src_range.Text))
                    {
                        MessageBox.Show("Select the Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        TB_src_range.Focus();
                        return;
                    }

                    else if (IsValidExcelCellReference(TB_src_range.Text) == false)
                    {
                        MessageBox.Show("Please Enter valid Source Range", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        TB_src_range.Focus();
                        return;
                    }

                    else if ((ax ?? "") != (workSheet2.Name ?? ""))
                    {
                        MessageBox.Show("Please select the range of the same worksheet", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        TB_src_range.Focus();
                        return;
                    }
                    else
                    {
                        goto GotoExpression;
                    }
                }
                else
                {
GotoExpression:
                    ;

                    var stringItems = new List<string>();

                    foreach (object item in List_Preview.Items)
                        stringItems.Add(item.ToString());

                    // Join the string representations into a single string
                    string items = string.Join(", ", stringItems);

                    des_rng.Validation.Delete();

                    // Create a new validation rule
                    var validation = des_rng.Validation;

                    // Add a drop-down list validation rule
                    validation.Delete();
                    validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, items, Type.Missing);
                    validation.IgnoreBlank = true;
                    validation.InCellDropdown = true;

                    des_rng.Select();
                    des_rng.set_Value(value: null);

                    Close();
                }
            }
            catch (Exception ex)
            {
                Close();
            }
        }

        private void Btn_Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Selection_Source_Click(object sender, EventArgs e)
        {
            if (selectedRange is null)
            {
                TB_src_range.Focus();
            }
            else
            {
                // TB_src_range.Text = selectedRange.Address


                Hide();

                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;

                Excel.Range userInput = (Excel.Range)excelApp.InputBox("Select a range", "Select a range", "=$A$1", Type: 8);
                src_rng = userInput;

                string sheetName;
                sheetName = Strings.Split(src_rng.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true), "]")[1];
                sheetName = Strings.Split(sheetName, "!")[0];

                if (Strings.Mid(sheetName, Strings.Len(sheetName), 1) == "'")
                {
                    sheetName = Strings.Mid(sheetName, 1, Strings.Len(sheetName) - 1);
                }

                workSheet = (Excel.Worksheet)workBook.Worksheets[sheetName];
                workSheet.Activate();

                src_rng.Select();
                // MsgBox(src_rng.Address)

                TB_src_range.Text = src_rng.get_Address();

                Show();
                TB_src_range.Focus();
                TB_src_range.Focus();

                // Define the range of cells to read (for example, cells A1 to A10)
                var range = src_rng;

                // Clear the ListBox
                List_Preview.Items.Clear();

                // Iterate over each cell in the range
                foreach (Excel.Range cell in range)
                {
                    // Add the cell's value to the ListBox
                    if (cell.get_Value() is not null)
                    {
                        List_Preview.Items.Add(cell.get_Value());
                    }
                }

                Label7.Visible = true;
                Label7.Text = List_Preview.Items.Count.ToString();
                TB_src_range.Focus();
                TB_src_range.Focus();
                // Me.Activate()

            }

        }


        private void Selection_Click(object sender, EventArgs e)
        {
            try
            {
                if (selectedRange is null)
                {
                    TB_dest_range.Focus();
                }
                else
                {

                    TB_dest_range.Text = selectedRange.get_Address();


                    // FocusedTextBox = 1
                    Hide();

                    excelApp = Globals.ThisAddIn.Application;
                    workBook = excelApp.ActiveWorkbook;

                    // Dim userInput As String = excelApp.InputBox("Select a range", "Select range", "=$A$1")


                    Excel.Range userInput = (Excel.Range)excelApp.InputBox("Select a range", "Select a range", "=$A$1", Type: 8);
                    des_rng = userInput;

                    string sheetName;
                    sheetName = Strings.Split(des_rng.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true), "]")[1];
                    sheetName = Strings.Split(sheetName, "!")[0];

                    if (Strings.Mid(sheetName, Strings.Len(sheetName), 1) == "'")
                    {
                        sheetName = Strings.Mid(sheetName, 1, Strings.Len(sheetName) - 1);
                    }

                    workSheet = (Excel.Worksheet)workBook.Worksheets[sheetName];
                    workSheet.Activate();

                    des_rng.Select();

                    TB_dest_range.Text = des_rng.get_Address();

                    Show();
                    TB_dest_range.Focus();
                    TB_dest_range.Focus();
                }
            }

            catch (Exception ex)
            {

                Show();
                TB_dest_range.Focus();

            }
        }



        private void Form1_Load(object sender, EventArgs e)
        {

            try
            {

                excelApp = Globals.ThisAddIn.Application;

                excelApp.SheetSelectionChange += excelApp_SheetSelectionChange;

                opened = opened + 1;

                if (excelApp.Selection is not null)
                {
                    selectedRange = (Excel.Range)excelApp.Selection;
                    des_rng = selectedRange;
                    TB_dest_range.Text = selectedRange.get_Address();
                }

                if (RadioButton1.Checked == true)
                {
                    ComboBox1.Enabled = false;
                    ListBox1.Enabled = false;
                    TB_src_range.Enabled = true;
                    Selection_Source.Enabled = true;
                }
            }

            catch (Exception ex)
            {

            }

        }

        private void ListBox1_DrawItem(object sender, DrawItemEventArgs e)
        {
            // If the index is invalid, exit
            if (e.Index < 0)
                return;
            Color backColor;

            // Determine the color based on even or odd index
            if (e.Index % 2 == 0)
            {
                // Odd lines
                e.Graphics.FillRectangle(Brushes.White, e.Bounds);
                backColor = Color.White;
            }
            else
            {
                // Even lines
                e.Graphics.FillRectangle(Brushes.LightGray, e.Bounds);
                backColor = Color.LightGray;
            }

            var textColor = Color.Black;


            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
            {
                // If item is selected, we'll use system colors to highlight.
                backColor = SystemColors.Highlight;
                textColor = SystemColors.HighlightText;
            }

            // Draw the text
            // e.Graphics.DrawString(ListBox1.Items(e.Index).ToString(), e.Font, Brushes.Black, e.Bounds)
            using (var brush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(brush, e.Bounds);
            }

            using (var brush = new SolidBrush(textColor))
            {
                e.Graphics.DrawString(ListBox1.Items[e.Index].ToString(), e.Font, brush, e.Bounds.Left, e.Bounds.Top);
            }

            // If the ListBox has focus, draw a focus rectangle around the selected item.
            e.DrawFocusRectangle();

        }

        // Private Sub ListBox12_DrawItem(sender As Object, e As DrawItemEventArgs) Handles ListBox1.DrawItem
        // If e.Index < 0 Then Return

        // '  Dim item As ColoredItem = CType(ListBox1.Items(e.Index), ColoredItem)

        // Dim textColor As Color = Color.Black
        // Dim backColor As Color = Color.White

        // If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
        // ' If item is selected, we'll use system colors to highlight.
        // backColor = SystemColors.Highlight
        // textColor = SystemColors.HighlightText
        // End If

        // ' Use the determined colors.
        // Using brush As New SolidBrush(backColor)
        // e.Graphics.FillRectangle(brush, e.Bounds)
        // End Using

        // ' Draw the text in the determined text color.
        // Using brush As New SolidBrush(textColor)
        // e.Graphics.DrawString(ListBox1.Items(e.Index).ToString(), e.Font, brush, e.Bounds.Left, e.Bounds.Top)
        // End Using

        // e.DrawFocusRectangle()
        // End Sub


        private void excelApp_SheetSelectionChange(object Sh, Excel.Range selectionRange1)
        {
            try
            {

                excelApp = Globals.ThisAddIn.Application;

                // If Me.ActiveControl Is TB_dest_range Then
                if (focuschange == false)
                {
                    if (TB_dest_range.Focused == true | ReferenceEquals(ActiveControl, TB_dest_range))
                    {
                        if (TB_dest_range.Focused == true)
                        {
                            des_rng = selectionRange1;
                        }
                        Activate();
                        BeginInvoke(new System.Action(() =>
                            {
                                TB_dest_range.Text = des_rng.get_Address();
                                SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                            }));
                    }

                    // ElseIf Me.ActiveControl Is TB_src_range Then
                    else if (TB_src_range.Focused == true | ReferenceEquals(ActiveControl, TB_src_range))
                    {
                        if (TB_src_range.Focused == true)
                        {
                            src_rng = selectionRange1;
                        }
                        Activate();
                        BeginInvoke(new System.Action(() =>
                            {
                                TB_src_range.Text = src_rng.get_Address();
                                SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                            }));

                    }
                }
            }



            catch (Exception ex)
            {

            }

        }

        private void TB_src_range_TextChanged(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            workBook = excelApp.ActiveWorkbook;
            workSheet = (Excel.Worksheet)workBook.ActiveSheet;

            try
            {

                if (TB_src_range.Text is not null & IsValidExcelCellReference(TB_src_range.Text) == true)
                {
                    focuschange = true;

                    // src_rng = excelApp.Range(cellAddress)
                    try
                    {
                        src_rng = excelApp.get_Range(TB_src_range.Text);
                        src_rng.Select();
                    }
                    catch
                    {
                        // Split the string into sheet name and cell address
                        string[] parts = TB_src_range.Text.Split('!');
                        string sheetName = parts[0];
                        string cellAddress = parts[1];

                        src_rng = excelApp.get_Range(cellAddress);
                        src_rng.Select();
                    }
                    // Define the range of cells to read (for example, cells A1 to A10)
                    if ((workSheet2.Name ?? "") != (workSheet.Name ?? ""))
                    {
                        TB_src_range.Text = workSheet.Name + "!" + src_rng.get_Address();
                        // src_rng = excelApp.Range(TB_src_range.Text)


                    }

                    var range = src_rng;

                    // Clear the ListBox
                    List_Preview.Items.Clear();

                    // Iterate over each cell in the range
                    foreach (Excel.Range cell in range)
                    {
                        // Add the cell's value to the ListBox
                        if (cell.get_Value() is not null)
                        {
                            List_Preview.Items.Add(cell.get_Value());
                        }
                    }

                    Label7.Visible = true;
                    Label7.Text = List_Preview.Items.Count.ToString();
                    Activate();
                    // TB_src_range.Focus()
                    TB_src_range.SelectionStart = TB_src_range.Text.Length;
                    focuschange = false;

                    ax = workSheet.Name;

                }
            }

            catch (Exception ex)
            {
                ax = "";
            }
        }

        private void TB_dest_rane_TextChanged(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            workBook = excelApp.ActiveWorkbook;
            workSheet = (Excel.Worksheet)workBook.ActiveSheet;
            try
            {

                if (TB_dest_range.Text is not null & IsValidExcelCellReference(TB_dest_range.Text) == true)
                {
                    focuschange = true;

                    // Define the range of cells to read (for example, cells A1 to A10)
                    des_rng = excelApp.get_Range(TB_dest_range.Text);
                    des_rng.Select();
                    var range = des_rng;

                    // Clear the ListBox
                    // List_Preview.Items.Clear()

                    // ' Iterate over each cell in the range
                    // For Each cell As Excel.Range In range
                    // ' Add the cell's value to the ListBox
                    // If cell.Value IsNot Nothing Then
                    // List_Preview.Items.Add(cell.Value)
                    // End If
                    // Next

                    // Label7.Visible = True
                    // Label7.Text = List_Preview.Items.Count
                    Activate();
                    // TB_src_range.Focus()
                    TB_dest_range.SelectionStart = TB_dest_range.Text.Length;
                    focuschange = false;
                    workSheet2 = workSheet;

                }
            }

            catch (Exception ex)
            {

            }
        }


        private void form(object sender, KeyEventArgs e)
        {

            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Btn_OK_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }

        }

        private void Listbox(object sender, KeyEventArgs e)
        {

            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Btn_OK_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }

        }

        private void Listboxx2(object sender, KeyEventArgs e)
        {

            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Btn_OK_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }

        }


        private void destination(object sender, KeyEventArgs e)
        {

            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Btn_OK_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }

        }

        private void source(object sender, KeyEventArgs e)
        {

            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Btn_OK_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }

        }

        private void TB_dest(object sender, KeyEventArgs e)
        {

            // Try
            // If e.KeyCode = Keys.Enter Then

            // Call Btn_OK_Click(sender, e)

            // End If

            // Catch ex As Exception

            // End Try

        }


        private void RB_1(object sender, KeyEventArgs e)
        {

            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Btn_OK_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }

        }

        private void RB_2(object sender, KeyEventArgs e)
        {

            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Btn_OK_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }

        }

        private void RB_3(object sender, KeyEventArgs e)
        {

            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Btn_OK_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }

        }



        private bool IsValidExcelCellReference(string cellReference)
        {

            // Regular expression pattern for a valid sheet name. This is a simplified version and might not cover all edge cases.
            // Excel sheet names cannot contain the characters \, /, *, [, ], :, ?, and cannot be 'History'.
            string sheetNamePattern = @"(?i)(?![\/*[\]:?])(?!History)[^\/\[\]*?:\\]+";

            // Regular expression pattern for a cell reference.
            // This pattern will match references like A1, $A$1, etc.
            string cellPattern = @"(\$?[A-Z]+\$?[0-9]+)";

            // Regular expression pattern for an Excel reference.
            // This pattern will match references like A1:B13, $A$1:$B$13, A1, $B$1, etc.
            string singleReferencePattern = cellPattern + "(:" + cellPattern + ")?";

            // Regular expression pattern to allow the sheet name, followed by '!', before the cell reference
            string fullPattern = "^(" + sheetNamePattern + "!)?(" + singleReferencePattern + ")(," + singleReferencePattern + ")*$";

            // Create a regex object with the pattern.
            var regex = new Regex(fullPattern);

            // Test the input string against the regex pattern.
            return regex.IsMatch(cellReference.ToUpper());

        }

        private void RadioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButton3.Checked == true)
            {
                ComboBox1.Enabled = true;
                ListBox1.Enabled = false;
                TB_src_range.Enabled = false;
                Selection_Source.Enabled = false;

            }
        }

        private void RadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButton2.Checked == true)
            {
                ComboBox1.Enabled = false;
                ListBox1.Enabled = true;
                TB_src_range.Enabled = false;
                Selection_Source.Enabled = false;
            }
        }

        private void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButton1.Checked == true)
            {
                ComboBox1.Enabled = false;
                ListBox1.Enabled = false;
                TB_src_range.Enabled = true;
                Selection_Source.Enabled = true;
            }
        }



        private void TB_dest_range_Enter(object sender, KeyEventArgs e)
        {
            // 'If Enter key is pressed then check if the text is a valid address
            // If IsValidExcelCellReference(TB_dest_range.Text) = True And e.KeyCode = Keys.Enter Then
            // des_rng = excelApp.Range(TB_dest_range.Text)
            // TB_dest_range.Focus()
            // des_rng.Select()

            // Call Btn_OK_Click(sender, e)   'OK button click event called

            // MsgBox(des_rng.Address)
            // ElseIf IsValidExcelCellReference(TB_dest_range.Text) = False And e.KeyCode = Keys.Enter Then
            // MessageBox.Show("Please Enter valid Destination Range", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            // TB_dest_range.Text = ""
            // TB_dest_range.Focus()
            // 'Me.Close()
            // Exit Sub
            // End If
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Btn_OK_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void TB_src_range_Enter(object sender, KeyEventArgs e)
        {
            // If Enter key is pressed then check if the text is a valid address

            // If IsValidExcelCellReference(TB_src_range.Text) = True And e.KeyCode = Keys.Enter Then
            // src_rng = excelApp.Range(TB_src_range.Text)
            // TB_src_range.Focus()
            // src_rng.Select()

            // Call Btn_OK_Click(sender, e)   'OK button click event called

            // 'MsgBox(des_rng.Address)
            // ElseIf IsValidExcelCellReference(TB_src_range.Text) = False And e.KeyCode = Keys.Enter Then
            // MessageBox.Show("Please Enter valid Source Range", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            // TB_src_range.Text = ""
            // TB_src_range.Focus()
            // 'Me.Close()
            // Exit Sub
            // End If
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Btn_OK_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void ComboBox1_TextUpdate(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(ComboBox1.Text))
            {
            }
            // Do nothing
            else
            {
                // Clear the list box
                List_Preview.Items.Clear();
                string selectedItem = ComboBox1.Text;

                // Check if the text has two consecutive commas.
                if (ComboBox1.Text.Contains(",,"))
                {
                    // Display a message to the user.
                    MessageBox.Show("Consecutive commas are not allowed.");

                    // Remove the last comma entered to prevent consecutive commas.
                    // Set the cursor at the end of the current text.
                    ComboBox1.Text = ComboBox1.Text.Remove(ComboBox1.Text.LastIndexOf(","), 1);
                    ComboBox1.SelectionStart = ComboBox1.Text.Length;
                    // Me.Refresh()

                    // Split the string into an array of strings
                    string[] items = ComboBox1.Text.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    for (int i = 0, loopTo = items.Length - 1; i <= loopTo; i++)


                        items[i] = items[i].TrimStart();


                    // ComboBox1.Items.AddRange(items)
                    List_Preview.Items.AddRange(items);
                    Label7.Visible = true;
                    Label7.Text = items.Count().ToString();
                }
                else
                {

                    // Split the string into an array of strings
                    // Dim items As String() = selectedItem.Split(","c)
                    string[] items = selectedItem.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    for (int i = 0, loopTo1 = items.Length - 1; i <= loopTo1; i++)
                        items[i] = items[i].TrimStart();


                    // ComboBox1.Items.AddRange(items)
                    List_Preview.Items.AddRange(items);
                    Label7.Visible = true;
                    Label7.Text = items.Count().ToString();

                }
            }
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form29_Simple_Drop_down_List_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form29_Simple_Drop_down_List_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form29_Simple_Drop_down_List_Shown(object sender, EventArgs e)
        {
            Focus();
            BringToFront();
            Activate();
            BeginInvoke(new System.Action(() =>
                {
                    TB_dest_range.Text = des_rng.get_Address();
                    SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                }));
        }

        private void Selection(object sender, EventArgs e)
        {

        }

        private void List_Preview_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void ComboBox2_MouseLeave(object sender, EventArgs e)
        {

        }

        private void ComboBox2_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Btn_OK_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void Selection(object sender, EventArgs e) => Selection();
    }

    public class ColoredItem1
    {
        public string Text { get; set; }
        public Color Color { get; set; } = Color.White;

        public ColoredItem1(string t)
        {
            Text = t;
        }

        public ColoredItem1(string t, Color c)
        {
            Text = t;
            Color = c;
        }

        public override string ToString()
        {
            return Text;
        }
    }
}