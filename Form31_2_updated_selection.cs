using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace VSTO_Addins
{

    public partial class Form31_2_updated_selection
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
        private Excel.Worksheet workSheet2;
        private Form31_UpdateDynamicDropdownList Form;

        public Form31_2_updated_selection()
        {
            InitializeComponent();
        }

        private void Form31_2_updated_selection_Load(object sender, EventArgs e)
        {

            PopulateDataGridViewWithExcelData();

            // Set only the non-checkbox columns to read-only
            foreach (DataGridViewColumn column in DataGridView1.Columns)
            {
                if (!(column is DataGridViewCheckBoxColumn))
                {
                    column.ReadOnly = true;
                }
            }

        }

        private DataGridViewTextBoxColumn CreateTextBoxColumn(string bindingName, string headerText)
        {
            var column = new DataGridViewTextBoxColumn();
            column.DataPropertyName = bindingName; // This should match the property name of the data you're binding to
            column.HeaderText = headerText;
            column.Name = bindingName;
            return column;
        }

        // This subroutine would be called to populate your DataGridView, assuming it's named dataGridView1.
        private void PopulateDataGridViewWithExcelData()
        {
            excelApp = Globals.ThisAddIn.Application;
            workBook = excelApp.ActiveWorkbook;
            workSheet = (Excel.Worksheet)workBook.ActiveSheet;
            try
            {
                // Create a new DataTable.
                var dataTable = new DataTable();

                // Define columns for the DataTable to match your Excel structure.
                // dataTable.Columns.Add("DataRange", GetType(String))
                dataTable.Columns.Add("OriginalDataRange", typeof(string));
                dataTable.Columns.Add("OutputRange", typeof(string));
                dataTable.Columns.Add("Level", typeof(int));
                dataTable.Columns.Add("Select", typeof(bool)); // For the CheckBox

                Excel.Worksheet targetWorksheet = null;
                foreach (Excel.Worksheet ws in excelApp.Worksheets)
                {
                    if (ws.Name == "MySpecialSheet")
                    {
                        targetWorksheet = ws;
                        break;
                    }
                }

                short Label;



                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(targetWorksheet.get_Range("A1").get_Value(), "", false)))
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(targetWorksheet.get_Range("A7").get_Value(), true, false)))
                    {
                        Label = 5; // Replace with the label you want if the condition is true
                    }
                    else
                    {
                        Label = 2;
                    } // Replace with the label you want if the condition is false

                    dataTable.Rows.Add(targetWorksheet.get_Range("A1").get_Value(), targetWorksheet.get_Range("A2").get_Value(), (object)Label, (object)false);

                }

                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(targetWorksheet.get_Range("B1").get_Value(), "", false)))
                {

                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(targetWorksheet.get_Range("B7").get_Value(), true, false)))
                    {
                        Label = 5; // Replace with the label you want if the condition is true
                    }
                    else
                    {
                        Label = 2;
                    } // Replace with the label you want if the condition is false

                    dataTable.Rows.Add(targetWorksheet.get_Range("B1").get_Value(), targetWorksheet.get_Range("B2").get_Value(), (object)Label, (object)false);

                }

                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(targetWorksheet.get_Range("C1").get_Value(), "", false)))
                {

                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(targetWorksheet.get_Range("C7").get_Value(), true, false)))
                    {
                        Label = 5; // Replace with the label you want if the condition is true
                    }
                    else
                    {
                        Label = 2;
                    } // Replace with the label you want if the condition is false

                    dataTable.Rows.Add(targetWorksheet.get_Range("C1").get_Value(), targetWorksheet.get_Range("C2").get_Value(), (object)Label, (object)false);

                }

                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(targetWorksheet.get_Range("D1").get_Value(), "", false)))
                {

                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(targetWorksheet.get_Range("D7").get_Value(), true, false)))
                    {
                        Label = 5; // Replace with the label you want if the condition is true
                    }
                    else
                    {
                        Label = 2;
                    } // Replace with the label you want if the condition is false

                    dataTable.Rows.Add(targetWorksheet.get_Range("D1").get_Value(), targetWorksheet.get_Range("D2").get_Value(), (object)Label, (object)false);

                }
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(targetWorksheet.get_Range("E1").get_Value(), "", false)))
                {

                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(targetWorksheet.get_Range("E7").get_Value(), true, false)))
                    {
                        Label = 5; // Replace with the label you want if the condition is true
                    }
                    else
                    {
                        Label = 2;
                    } // Replace with the label you want if the condition is false

                    dataTable.Rows.Add(targetWorksheet.get_Range("E1").get_Value(), targetWorksheet.get_Range("E2").get_Value(), (object)Label, (object)false);

                }

                // Set the DataGridView's DataSource to the DataTable.
                DataGridView1.DataSource = dataTable;

                // Adjusting the DataGridView properties for better appearance
                DataGridView1.AutoResizeColumns();
                DataGridView1.Columns["Select"].DisplayIndex = 0; // To show the checkbox column as the first column
            }
            catch (Exception ex)
            {
                Interaction.MsgBox("Dynamic Drop-down List is not available");
            }
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            // If the clicked cell is in the checkbox column.
            if (e.ColumnIndex == DataGridView1.Columns["Select"].Index)
            {
                DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }

        }
        // The event handler for CellValueChanged
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // Check if the change happened in the checkbox column
            if (e.ColumnIndex == DataGridView1.Columns["Select"].Index)
            {
                UpdateRowColor(e.RowIndex);
            }

        }

        // Call this method in the CellValueChanged event handler to update the row color.
        private void UpdateRowColor(int rowIndex)
        {
            if (rowIndex < 0)
                return;

            var row = DataGridView1.Rows[rowIndex];
            bool isChecked = Convert.ToBoolean(row.Cells["Select"].Value);

            if (isChecked)
            {
                row.DefaultCellStyle.BackColor = SystemColors.Highlight;
                row.DefaultCellStyle.ForeColor = Color.White;
                row.DefaultCellStyle.Font = new Font("Segoe UI", 10f);
            }
            // row.DefaultCellStyle.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
            else
            {
                row.DefaultCellStyle.BackColor = Color.White;
                row.DefaultCellStyle.ForeColor = Color.Black;
                row.DefaultCellStyle.Font = new Font("Segoe UI", 10f);
                // row.DefaultCellStyle.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Regular)
            }
        }

        // Handle the event when the DataGridView data binding is complete to color initial rows (if necessary)
        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {

            foreach (DataGridViewRow row in DataGridView1.Rows)
                UpdateRowColor(row.Index);

        }

        // Handle the CellClick event for the DataGridView.
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            // Ignore header clicks or clicks on the checkbox cell itself
            if (e.RowIndex < 0 | e.ColumnIndex == DataGridView1.Columns["Select"].Index)
                return;

            // Get the checkbox cell
            DataGridViewCheckBoxCell checkBoxCell = DataGridView1.Rows[e.RowIndex].Cells["Select"] as DataGridViewCheckBoxCell;

            if (checkBoxCell is not null && !checkBoxCell.ReadOnly)
            {
                // Toggle the checkbox value
                checkBoxCell.Value = !Convert.ToBoolean(checkBoxCell.Value);
                // Commit the edit immediately
                DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }

        }

        private void Btn_OK_Click(object sender, EventArgs e)
        {

            var targetWorksheet = default(Excel.Worksheet);
            int i = 1;
            foreach (var ws in excelApp.ActiveWorkbook.Worksheets)
            {
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.name, "MySpecialSheet", false)))
                {
                    targetWorksheet = (Excel.Worksheet)ws;
                    break;
                }
            }

            // This list will hold the values of the checked rows
            var checkedRowsValues = new List<string>();

            // Iterate over each row to check if the checkbox is checked
            foreach (DataGridViewRow row in DataGridView1.Rows)
            {
                bool isSelected = Convert.ToBoolean(row.Cells["Select"].Value); // Replace "Select" with your checkbox column's name
                if (isSelected & i == 1)
                {

                    GlobalModule.Variable1 = targetWorksheet.get_Range("A1").get_Value().ToString();
                    GlobalModule.Variable2 = targetWorksheet.get_Range("A2").get_Value().ToString();
                    GlobalModule.Header = Conversions.ToBoolean(targetWorksheet.get_Range("A3").get_Value().ToString());
                    GlobalModule.Ascending = Conversions.ToBoolean(targetWorksheet.get_Range("A4").get_Value().ToString());
                    GlobalModule.Descending = Conversions.ToBoolean(targetWorksheet.get_Range("A5").get_Value().ToString());
                    GlobalModule.TextConvert = Conversions.ToBoolean(targetWorksheet.get_Range("A6").get_Value().ToString());
                    GlobalModule.OptionType = Conversions.ToBoolean(targetWorksheet.get_Range("A7").get_Value().ToString());
                    GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetWorksheet.get_Range("A8").get_Value().ToString());
                    GlobalModule.Flag_CreateDDDL = Conversions.ToBoolean(targetWorksheet.get_Range("A9").get_Value().ToString());
                    GlobalModule.sheetName10 = targetWorksheet.get_Range("A10").get_Value().ToString();
                    GlobalModule.sheetName11 = targetWorksheet.get_Range("A11").get_Value().ToString();
                    Form = new Form31_UpdateDynamicDropdownList();
                    Form.Show();
                    Form.TextBox1.Text = i.ToString();
                }

                else if (isSelected & i == 2)
                {

                    GlobalModule.Variable1 = targetWorksheet.get_Range("B1").get_Value().ToString();
                    GlobalModule.Variable2 = targetWorksheet.get_Range("B2").get_Value().ToString();
                    GlobalModule.Header = Conversions.ToBoolean(targetWorksheet.get_Range("B3").get_Value().ToString());
                    GlobalModule.Ascending = Conversions.ToBoolean(targetWorksheet.get_Range("B4").get_Value().ToString());
                    GlobalModule.Descending = Conversions.ToBoolean(targetWorksheet.get_Range("B5").get_Value().ToString());
                    GlobalModule.TextConvert = Conversions.ToBoolean(targetWorksheet.get_Range("B6").get_Value().ToString());
                    GlobalModule.OptionType = Conversions.ToBoolean(targetWorksheet.get_Range("B7").get_Value().ToString());
                    GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetWorksheet.get_Range("B8").get_Value().ToString());
                    GlobalModule.Flag_CreateDDDL = Conversions.ToBoolean(targetWorksheet.get_Range("B9").get_Value().ToString());
                    GlobalModule.sheetName10 = targetWorksheet.get_Range("B10").get_Value().ToString();
                    GlobalModule.sheetName11 = targetWorksheet.get_Range("B11").get_Value().ToString();
                    Form = new Form31_UpdateDynamicDropdownList();
                    Form.Show();
                    Form.TextBox1.Text = i.ToString();
                }

                else if (isSelected & i == 3)
                {
                    GlobalModule.Variable1 = targetWorksheet.get_Range("C1").get_Value().ToString();
                    GlobalModule.Variable2 = targetWorksheet.get_Range("C2").get_Value().ToString();
                    GlobalModule.Header = Conversions.ToBoolean(targetWorksheet.get_Range("C3").get_Value().ToString());
                    GlobalModule.Ascending = Conversions.ToBoolean(targetWorksheet.get_Range("C4").get_Value().ToString());
                    GlobalModule.Descending = Conversions.ToBoolean(targetWorksheet.get_Range("C5").get_Value().ToString());
                    GlobalModule.TextConvert = Conversions.ToBoolean(targetWorksheet.get_Range("C6").get_Value().ToString());
                    GlobalModule.OptionType = Conversions.ToBoolean(targetWorksheet.get_Range("C7").get_Value().ToString());
                    GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetWorksheet.get_Range("C8").get_Value().ToString());
                    GlobalModule.Flag_CreateDDDL = Conversions.ToBoolean(targetWorksheet.get_Range("C9").get_Value().ToString());
                    GlobalModule.sheetName10 = targetWorksheet.get_Range("C10").get_Value().ToString();
                    GlobalModule.sheetName11 = targetWorksheet.get_Range("C11").get_Value().ToString();
                    Form = new Form31_UpdateDynamicDropdownList();
                    Form.Show();
                    Form.TextBox1.Text = i.ToString();
                }

                else if (isSelected & i == 4)
                {

                    GlobalModule.Variable1 = targetWorksheet.get_Range("D1").get_Value().ToString();
                    GlobalModule.Variable2 = targetWorksheet.get_Range("D2").get_Value().ToString();
                    GlobalModule.Header = Conversions.ToBoolean(targetWorksheet.get_Range("D3").get_Value().ToString());
                    GlobalModule.Ascending = Conversions.ToBoolean(targetWorksheet.get_Range("D4").get_Value().ToString());
                    GlobalModule.Descending = Conversions.ToBoolean(targetWorksheet.get_Range("D5").get_Value().ToString());
                    GlobalModule.TextConvert = Conversions.ToBoolean(targetWorksheet.get_Range("D6").get_Value().ToString());
                    GlobalModule.OptionType = Conversions.ToBoolean(targetWorksheet.get_Range("D7").get_Value().ToString());
                    GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetWorksheet.get_Range("D8").get_Value().ToString());
                    GlobalModule.Flag_CreateDDDL = Conversions.ToBoolean(targetWorksheet.get_Range("D9").get_Value().ToString());
                    GlobalModule.sheetName10 = targetWorksheet.get_Range("D10").get_Value().ToString();
                    GlobalModule.sheetName11 = targetWorksheet.get_Range("D11").get_Value().ToString();
                    Form = new Form31_UpdateDynamicDropdownList();
                    Form.Show();
                    Form.TextBox1.Text = i.ToString();
                }

                else if (isSelected & i == 5)
                {
                    GlobalModule.Variable1 = targetWorksheet.get_Range("E1").get_Value().ToString();
                    GlobalModule.Variable2 = targetWorksheet.get_Range("E2").get_Value().ToString();
                    GlobalModule.Header = Conversions.ToBoolean(targetWorksheet.get_Range("E3").get_Value().ToString());
                    GlobalModule.Ascending = Conversions.ToBoolean(targetWorksheet.get_Range("E4").get_Value().ToString());
                    GlobalModule.Descending = Conversions.ToBoolean(targetWorksheet.get_Range("E5").get_Value().ToString());
                    GlobalModule.TextConvert = Conversions.ToBoolean(targetWorksheet.get_Range("E6").get_Value().ToString());
                    GlobalModule.OptionType = Conversions.ToBoolean(targetWorksheet.get_Range("E7").get_Value().ToString());
                    GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetWorksheet.get_Range("E8").get_Value().ToString());
                    GlobalModule.Flag_CreateDDDL = Conversions.ToBoolean(targetWorksheet.get_Range("E9").get_Value().ToString());
                    GlobalModule.sheetName10 = targetWorksheet.get_Range("E10").get_Value().ToString();
                    GlobalModule.sheetName11 = targetWorksheet.get_Range("E11").get_Value().ToString();
                    Form = new Form31_UpdateDynamicDropdownList();
                    Form.Show();
                    Form.TextBox1.Text = i.ToString();

                }
                i = i + 1;

            }
            Close();

        }

        private void Btn_Cancel_Click(object sender, EventArgs e)
        {
            Dispose();
        }

    }
}