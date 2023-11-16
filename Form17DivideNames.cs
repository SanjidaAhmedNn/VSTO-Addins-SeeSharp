using System;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace VSTO_Addins
{

    public partial class Form17DivideNames
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
        private Range sourceRange, destRange;
        private Range selectedRange;
        private string[] mainArr = new string[7];
        private bool changeState = false;
        private bool txtChanged = false;

        public Form17DivideNames()
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

        private void Form17DivideNames_Load(object sender, EventArgs e)
        {

            try
            {

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selectedRng = (Range)excelApp.Selection;
                txtSourceRange.Text = selectedRng.get_Address();
                RB_Same_As_Source_Range.Checked = true;

                KeyPreview = true;
            }

            catch (Exception ex)
            {

            }

        }


        private void txtSourceRange_TextChanged(object sender, EventArgs e)
        {

            try
            {

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;


                txtChanged = true;
                sourceRange = worksheet.get_Range(txtSourceRange.Text);


                sourceRange.Select();




                if (changeState == true)
                {


                    if ((destRange.Worksheet.Name ?? "") != (sourceRange.Worksheet.Name ?? ""))
                    {

                        txtDestRange.Text = destRange.Worksheet.Name + "!" + destRange.get_Address();

                    }


                }
            }



            catch (Exception ex)
            {

            }

            txtChanged = false;

            txtSourceRange.Focus();
            display();

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


                if ((destRange.Worksheet.Name ?? "") != (sourceRange.Worksheet.Name ?? ""))
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




                txtDestRange.Text = destRange.Worksheet.Name + "!" + destRange.get_Address();

                destRange.Select();
                txtDestRange.Focus();
            }




            catch (Exception ex)
            {

                txtDestRange.Focus();

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
                sourceRange = (Range)excelApp.InputBox("Please Select the First Range", "First Range Selection", selectedRange.get_Address(), Type: 8);
                Show();



                txtSourceRange.Text = sourceRange.Worksheet.Name + "!" + sourceRange.get_Address();

                sourceRange.Select();

                txtSourceRange.Focus();
            }



            catch (Exception ex)
            {

                txtSourceRange.Focus();

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

        public bool IsValidRng(string input)
        {
            // "^(([A-Za-z]+[0-9]*( \([0-9]+\))?!)?\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,([A-Za-z]+[0-9]*( \([0-9]+\))?!)?\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$"     
            string pattern = @"^(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$";
            return Regex.IsMatch(input, pattern);

        }

        private void btnOK_Click(object sender, EventArgs e)
        {

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                selectedRange = (Range)excelApp.Selection;

                int checkBox_checked_count = 0;
                foreach (Control ctrl in CustomGroupBox7.Controls)
                {
                    if (ctrl is System.Windows.Forms.CheckBox)
                    {
                        System.Windows.Forms.CheckBox chk = (System.Windows.Forms.CheckBox)ctrl;
                        if (chk.Checked)
                        {
                            checkBox_checked_count += 1;
                        }
                    }
                }


                if (RB_Same_As_Source_Range.Checked == true)
                {

                    if (string.IsNullOrEmpty(txtSourceRange.Text))
                    {
                        Interaction.MsgBox("Please select the Source Range.", MsgBoxStyle.Exclamation, "Error!");
                        txtSourceRange.Focus();
                        return;
                    }
                    else if (IsValidRng(txtSourceRange.Text.ToUpper()) == false)
                    {
                        Interaction.MsgBox("Please use a valid range in the Source Range.", MsgBoxStyle.Exclamation, "Error!");
                        txtDestRange.Focus();
                        return;
                    }


                    else if (checkBox_checked_count == 0)
                    {
                        Interaction.MsgBox("Please check least one checkbox to divide names.", MsgBoxStyle.Exclamation, "Error!");

                        CustomGroupBox7.Focus();
                        return;


                    }
                }

                else if (RB_Different_Range.Checked == true)
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
                            Interaction.MsgBox("Please select the Source Range.", MsgBoxStyle.Exclamation, "Error!");
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
                            Interaction.MsgBox("Please use a valid range in the Source Range.", MsgBoxStyle.Exclamation, "Error!");
                            txtSourceRange.Text = "";
                            txtSourceRange.Focus();
                            return;
                        }
                    }

                    else if (!string.IsNullOrEmpty(txtSourceRange.Text) & !string.IsNullOrEmpty(txtDestRange.Text))
                    {
                        if (checkBox_checked_count == 0)
                        {
                            Interaction.MsgBox("Please check least one checkbox to divide names.", MsgBoxStyle.Exclamation, "Error!");
                            CustomGroupBox7.Focus();
                            return;
                        }

                        else if (IsValidRng(txtSourceRange.Text.ToUpper()) == false & IsValidRng(txtDestRange.Text.ToUpper()) == true)
                        {
                            Interaction.MsgBox("Please use a valid range in the Source Range.", MsgBoxStyle.Exclamation, "Error!");
                            txtSourceRange.Text = "";
                            txtSourceRange.Focus();
                            return;
                        }

                        else if (IsValidRng(txtSourceRange.Text.ToUpper()) == true & IsValidRng(txtDestRange.Text.ToUpper()) == false)
                        {
                            Interaction.MsgBox("Please use a valid range in the Destination Range.", MsgBoxStyle.Exclamation, "Error!");
                            txtDestRange.Text = "";
                            txtDestRange.Focus();
                            return;
                        }
                        else if (IsValidRng(txtSourceRange.Text.ToUpper()) == false & IsValidRng(txtDestRange.Text.ToUpper()) == false)
                        {
                            Interaction.MsgBox("Please use valid ranges in the Source Range and in the Destination Range.", MsgBoxStyle.Exclamation, "Error!");
                            txtSourceRange.Text = "";
                            txtDestRange.Text = "";
                            txtSourceRange.Focus();
                            return;

                        }

                    }

                }


                string[] arrRng;
                string temp;
                temp = txtSourceRange.Text;
                worksheet1 = sourceRange.Worksheet;


                if (CB_Backup_Sheet.Checked == true)
                {

                    workbook.ActiveSheet.Copy(After: workbook.Sheets[workbook.Sheets.Count]);
                    outWorksheet = (Excel.Worksheet)workbook.Sheets[workbook.Sheets.Count];


                    worksheet1.Activate();
                    txtSourceRange.Text = temp;

                }



                if (RB_Same_As_Source_Range.Checked == true)
                {

                    var outputColumn = default(int);

                    arrRng = Strings.Split(txtSourceRange.Text, ",");

                    var headerIndex = default(int);
                    for (int i = 0, loopTo = Information.UBound(arrRng); i <= loopTo; i++)
                    {

                        sourceRange = worksheet.get_Range(arrRng[i]);

                        for (int j = 1, loopTo1 = sourceRange.Rows.Count; j <= loopTo1; j++)
                        {

                            mainArr = new[] { "", "", "", "", "", "", "", "" };
                            Name = Conversions.ToString(sourceRange.Cells[j, 1].value);

                            nameSplitter();
                            string headerStr = "";
                            foreach (Control ctrl in CustomGroupBox7.Controls)
                            {
                                if (ctrl is System.Windows.Forms.CheckBox)
                                {
                                    System.Windows.Forms.CheckBox chk = (System.Windows.Forms.CheckBox)ctrl;
                                    if (chk.Checked)
                                    {
                                        headerStr = headerStr + "," + chk.Text;
                                    }
                                }
                            }

                            headerStr = headerStr.Replace("Select All,", string.Empty);
                            headerStr = Strings.Right(headerStr, Strings.Len(headerStr) - 1);

                            string[] arrHeaderStr = Strings.Split(headerStr, ",");
                            outputColumn = Information.UBound(arrHeaderStr) + 1;
                            for (int k = 0, loopTo2 = Information.UBound(arrHeaderStr); k <= loopTo2; k++)
                            {
                                worksheet.Cells[100000, k + 1].value = arrHeaderStr[k];

                                switch (arrHeaderStr[k] ?? "")
                                {
                                    case var @case when @case == "Title":
                                        {
                                            headerIndex = 0;
                                            break;
                                        }
                                    case var case1 when case1 == "First Name":
                                        {
                                            headerIndex = 1;
                                            break;
                                        }
                                    case var case2 when case2 == "Middle Name":
                                        {
                                            headerIndex = 2;
                                            break;
                                        }
                                    case var case3 when case3 == "Last Name Prefix":
                                        {
                                            headerIndex = 3;
                                            break;
                                        }
                                    case var case4 when case4 == "Last Name":
                                        {
                                            headerIndex = 4;
                                            break;
                                        }
                                    case var case5 when case5 == "Name Suffix":
                                        {
                                            headerIndex = 5;
                                            break;
                                        }
                                    case var case6 when case6 == "Name Abbreviations":
                                        {
                                            headerIndex = 6;
                                            break;
                                        }

                                }
                                if (CB_Keep_Formatting.Checked == true)
                                {
                                    sourceRange.Cells[j, 1].copy(worksheet.Cells[100000 + j, k + 1]);
                                }
                                worksheet.Cells[100000 + j, k + 1].value = mainArr[headerIndex];
                            }





                        }

                        worksheet.get_Range(worksheet.Cells[100000, 1], worksheet.Cells[100000, outputColumn]).Font.Bold = (object)true;
                        worksheet.get_Range(worksheet.Cells[100000, 1], worksheet.Cells[100000 + sourceRange.Rows.Count, outputColumn]).Copy(sourceRange.Cells[1, 1]);

                        worksheet.get_Range(worksheet.Cells[100000, 1], worksheet.Cells[100000 + sourceRange.Rows.Count, outputColumn]).Delete(XlDeleteShiftDirection.xlShiftUp);
                        worksheet.get_Range(sourceRange.Cells[1, 1], sourceRange.Cells[sourceRange.Rows.Count + 1, outputColumn]).Select();

                        var border = selectedRange.Borders;
                        border.LineStyle = XlLineStyle.xlContinuous;
                        border.Weight = XlBorderWeight.xlThin;

                        selectedRange.EntireColumn.AutoFit();

                        if (CB_Add_Header.Checked == false)
                        {

                            worksheet.get_Range(selectedRange.Cells[1, 1], selectedRange.Cells[1, outputColumn]).Delete(XlDeleteShiftDirection.xlShiftUp);
                            worksheet.get_Range(sourceRange.Cells[1, 1], sourceRange.Cells[sourceRange.Rows.Count + 1, outputColumn]).Select();

                        }

                    }
                }



                else if (RB_Different_Range.Checked == true)
                {


                    var outputColumn = default(int);

                    arrRng = Strings.Split(txtSourceRange.Text, ",");

                    var headerIndex1 = default(int);
                    for (int i = 0, loopTo3 = Information.UBound(arrRng); i <= loopTo3; i++)
                    {

                        sourceRange = worksheet.get_Range(arrRng[i]);

                        for (int j = 1, loopTo4 = sourceRange.Rows.Count; j <= loopTo4; j++)
                        {

                            mainArr = new[] { "", "", "", "", "", "", "", "" };
                            Name = Conversions.ToString(sourceRange.Cells[j, 1].value);

                            nameSplitter();
                            string headerStr = "";
                            foreach (Control ctrl in CustomGroupBox7.Controls)
                            {
                                if (ctrl is System.Windows.Forms.CheckBox)
                                {
                                    System.Windows.Forms.CheckBox chk = (System.Windows.Forms.CheckBox)ctrl;
                                    if (chk.Checked)
                                    {
                                        headerStr = headerStr + "," + chk.Text;
                                    }
                                }
                            }

                            headerStr = headerStr.Replace("Select All,", string.Empty);
                            headerStr = Strings.Right(headerStr, Strings.Len(headerStr) - 1);

                            string[] arrHeaderStr = Strings.Split(headerStr, ",");
                            outputColumn = Information.UBound(arrHeaderStr) + 1;
                            for (int k = 0, loopTo5 = Information.UBound(arrHeaderStr); k <= loopTo5; k++)
                            {
                                worksheet.Cells[100000, k + 1].value = arrHeaderStr[k];

                                switch (arrHeaderStr[k] ?? "")
                                {
                                    case var case7 when case7 == "Title":
                                        {
                                            headerIndex1 = 0;
                                            break;
                                        }
                                    case var case8 when case8 == "First Name":
                                        {
                                            headerIndex1 = 1;
                                            break;
                                        }
                                    case var case9 when case9 == "Middle Name":
                                        {
                                            headerIndex1 = 2;
                                            break;
                                        }
                                    case var case10 when case10 == "Last Name Prefix":
                                        {
                                            headerIndex1 = 3;
                                            break;
                                        }
                                    case var case11 when case11 == "Last Name":
                                        {
                                            headerIndex1 = 4;
                                            break;
                                        }
                                    case var case12 when case12 == "Name Suffix":
                                        {
                                            headerIndex1 = 5;
                                            break;
                                        }
                                    case var case13 when case13 == "Name Abbreviations":
                                        {
                                            headerIndex1 = 6;
                                            break;
                                        }

                                }
                                if (CB_Keep_Formatting.Checked == true)
                                {
                                    sourceRange.Cells[j, 1].copy(worksheet.Cells[100000 + j, k + 1]);
                                }
                                worksheet.Cells[100000 + j, k + 1].value = mainArr[headerIndex1];
                            }





                        }

                        worksheet.get_Range(worksheet.Cells[100000, 1], worksheet.Cells[100000, outputColumn]).Font.Bold = (object)true;
                        worksheet.get_Range(worksheet.Cells[100000, 1], worksheet.Cells[100000 + sourceRange.Rows.Count, outputColumn]).Copy(destRange.Cells[1, 1]);

                        worksheet.get_Range(worksheet.Cells[100000, 1], worksheet.Cells[100000 + sourceRange.Rows.Count, outputColumn]).Delete(XlDeleteShiftDirection.xlShiftUp);
                        worksheet.get_Range(destRange.Cells[1, 1], destRange.Cells[sourceRange.Rows.Count + 1, outputColumn]).Select();

                        var border = selectedRange.Borders;
                        border.LineStyle = XlLineStyle.xlContinuous;
                        border.Weight = XlBorderWeight.xlThin;

                        selectedRange.EntireColumn.AutoFit();

                        if (CB_Add_Header.Checked == false)
                        {

                            worksheet.get_Range(selectedRange.Cells[1, 1], selectedRange.Cells[1, outputColumn]).Delete(XlDeleteShiftDirection.xlShiftUp);
                            worksheet.get_Range(destRange.Cells[1, 1], destRange.Cells[sourceRange.Rows.Count + 1, outputColumn]).Select();

                        }

                    }




                }


                Dispose();
            }




            catch (Exception ex)
            {

            }

            Dispose();



        }




        public bool checkTitle(string inputStr)
        {

            int dotCount;

            string[] arrTitle = new string[] { "Mr", "Mister", "Mrs", "Missus", "Miss", "Ms", "Dr", "Doctor", "Prof", "Professor", "Sir", "Lady", "Lord", "Madam", "Mdm", "Count", "Madame", "Master", "Rev", "Reverend", "Fr", "Father", "Sr", "Sister", "Pvt", "Private", "Esq", "Esquire", "Imam", "Sheikh", "Capt", "Captain", "Cpl", "Corporal", "Sgt", "Sergeant", "Gen", "General", "Lt", "Lieutenant", "Eng", "Engineer", "Hon", "Honorable", "Pres", "President", "VP", "Vice President", "Gov", "Governor", "Sen", "Senator", "Rep", "Representative", "Mx", "Herr", "Frau", "Duke", "Señor", "Señora", "Señorita", "Dott", "Dottore", "Mlle", "Mademoiselle", "Maestro", "Don", "Doña", "Smt", "Shrimati", "Shri", "Guru", "Sensei" };

            dotCount = 0;

            // count if there is any periods in the first word of the name
            // if a period is there then it will be considered as a title
            foreach (char c in inputStr)
            {

                if (Conversions.ToString(c) == ".")
                {
                    dotCount += 1;
                }

            }

            // checks if there are period(s) in the first word
            // OR the first word matches with any of the word from the arrTitle array (case insensitively)
            // if any one of the 2 conditon is true then, assign the first word as value in the first column of the destRange
            // otherwise assign a blank value
            if (dotCount > 0 | arrTitle.Contains(inputStr, StringComparer.OrdinalIgnoreCase))
            {
                return true;
            }
            else
            {
                return false;
            }

        }


        public bool checkSuffix(string inputStr)
        {

            int dotCount;

            string[] arrSuffix = new string[] { "Jr", "Sr", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "MD", "PhD", "Esq", "DDS", "RN", "CPA", "DVM", "JD", "LLB", "LLM", "BA", "BS", "MA", "MS", "PsyD", "OD", "DO", "EdD", "DPhil", "PE", "CFA", "MBA", "MPH", "BEd", "MFA", "ThD", "DMin", "DPT", "BBA", "MDiv", "RPh", "OBE", "KBE", "DC", "NP", "PA", "CNM", "FACP", "DABR" };

            // Name Suffix


            // count if there is any periods in the last word of the name
            // if a period is there then it will be considered as a Name Suffix
            dotCount = 0;
            foreach (char c in inputStr)
            {

                if (Conversions.ToString(c) == ".")
                {
                    dotCount += 1;
                }

            }

            // checks if there are period(s) in the last word
            // OR the last word matches with any of the word from the arrSuffix array (case insensitively)
            // if any one of the 2 conditon is true then, assign the last word as value in the last column of the destRange
            // otherwise assign a blank value
            if (dotCount > 0 | arrSuffix.Contains(inputStr, StringComparer.OrdinalIgnoreCase))
            {
                return true;
            }
            else
            {
                return false;
            }

        }




        public void nameSplitter()
        {

            excelApp = Globals.ThisAddIn.Application;
            workbook = excelApp.ActiveWorkbook;
            worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            selectedRange = (Range)excelApp.Selection;
            // Dim arrRng As String()
            string[] arrName;

            // Dim mainArr(7) As String


            string[] arrTitle = new string[] { "Mr", "Mister", "Mrs", "Missus", "Miss", "Ms", "Dr", "Doctor", "Prof", "Professor", "Sir", "Lady", "Lord", "Madam", "Mdm", "Count", "Madame", "Master", "Rev", "Reverend", "Fr", "Father", "Sr", "Sister", "Pvt", "Private", "Esq", "Esquire", "Imam", "Sheikh", "Capt", "Captain", "Cpl", "Corporal", "Sgt", "Sergeant", "Gen", "General", "Lt", "Lieutenant", "Eng", "Engineer", "Hon", "Honorable", "Pres", "President", "VP", "Vice President", "Gov", "Governor", "Sen", "Senator", "Rep", "Representative", "Mx", "Herr", "Frau", "Duke", "Señor", "Señora", "Señorita", "Dott", "Dottore", "Mlle", "Mademoiselle", "Maestro", "Don", "Doña", "Smt", "Shrimati", "Shri", "Guru", "Sensei" };


            string[] arrSuffix = new string[] { "Jr", "Sr", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "MD", "PhD", "Esq", "DDS", "RN", "CPA", "DVM", "JD", "LLB", "LLM", "BA", "BS", "MA", "MS", "PsyD", "OD", "DO", "EdD", "DPhil", "PE", "CFA", "MBA", "MPH", "BEd", "MFA", "ThD", "DMin", "DPT", "BBA", "MDiv", "RPh", "OBE", "KBE", "DC", "NP", "PA", "CNM", "FACP", "DABR" };


            // Dim arrHeader() = {"Full Name", "Title", "First Name", "Middle Name", "Last Name Prefix", "Last Name", "Abbreviations", "Name Suffix"}
            // Dim arrSplitName As String()


            // arrRng = Split(txtSourceRange.Text, ",")

            // For i = 0 To UBound(arrRng)

            // selectedRange = worksheet.Range(arrRng(i))

            // For j = 1 To selectedRange.Rows.Count

            // mainArr = {"", "", "", "", "", "", "", ""}
            // mainArr(0) = selectedRange.Cells(j, 1).value

            // arrName = Split(selectedRange.Cells(j, 1).value, " ")
            arrName = Strings.Split(Name, " ");



            if (Information.UBound(arrName) == 0)
            {

                if (checkTitle(arrName[0]) == true)
                {
                    mainArr[0] = arrName[0];
                }

                else if (checkSuffix(arrName[0]) == true)
                {
                    mainArr[5] = arrName[0];
                }

                else
                {
                    mainArr[1] = arrName[0];

                }
            }



            else if (Information.UBound(arrName) == 1)
            {

                if (checkTitle(arrName[0]) == true & checkSuffix(arrName[1]) == true)
                {
                    // Dr. PhD

                    // add title to 1st place in Mainarray
                    mainArr[0] = arrName[0];

                    // add suffix to 6th  place in main array
                    mainArr[5] = arrName[1];
                }

                else if (checkTitle(arrName[0]) == true & checkSuffix(arrName[1]) == false)
                {
                    // Dr. John

                    // add title in the 1st place
                    mainArr[0] = arrName[0];

                    // add first name to the 2nd place 
                    mainArr[1] = arrName[1];
                }

                else if (checkTitle(arrName[0]) == false & checkSuffix(arrName[1]) == true)
                {
                    // John PhD

                    // add first name to the 2nd place
                    mainArr[1] = arrName[0];

                    // add suffix to the 6th place
                    mainArr[6] = arrName[1];
                }

                else if (checkTitle(arrName[0]) == false & checkSuffix(arrName[1]) == false)
                {
                    // John Smith

                    // add first name to the 2nd place
                    mainArr[1] = arrName[0];

                    // add last name in the 5th place
                    mainArr[4] = arrName[1];

                }
            }


            else if (Information.UBound(arrName) == 2)
            {

                if (checkTitle(arrName[0]) == true & checkSuffix(arrName[2]) == true)
                {
                    // Dr. John PhD

                    // add title to 1st place in Mainarray
                    mainArr[0] = arrName[0];

                    // add first name to the 2nd plcae in the main array
                    mainArr[1] = arrName[1];

                    // add suffix to 6th place in main array
                    mainArr[5] = arrName[2];
                }


                else if (checkTitle(arrName[0]) == true & checkSuffix(arrName[2]) == false)
                {
                    // Dr. John Smith

                    // add title to the 1st place
                    mainArr[0] = arrName[0];

                    // add frist name to the 2nd place
                    mainArr[1] = arrName[1];

                    // add last name to the 5th place
                    mainArr[4] = arrName[2];
                }

                else if (checkTitle(arrName[0]) == false & checkSuffix(arrName[2]) == true)
                {
                    // John Smith PhD

                    // add first name to the 2nd place
                    mainArr[1] = arrName[0];

                    // add last name to the 5th place
                    mainArr[4] = arrName[1];

                    // add suffix to the 6th place
                    mainArr[5] = arrName[2];
                }

                else if (checkTitle(arrName[0]) == false & checkSuffix(arrName[2]) == false)
                {
                    // John Phillip Smith

                    // add first name to the 2nd place
                    mainArr[1] = arrName[0];

                    // add middle name to the 3rd place
                    mainArr[2] = arrName[1];

                    // add last name to the 5th plcae
                    mainArr[4] = arrName[2];

                    // add abbreviation to the 7th place
                    mainArr[6] = Strings.Left(mainArr[1], 1) + "." + Strings.Left(mainArr[2], 1) + ". " + mainArr[4];


                }
            }


            else if (Information.UBound(arrName) == 3)
            {


                if (checkTitle(arrName[0]) == true & checkSuffix(arrName[3]) == true)
                {
                    // Dr. John Smith PhD

                    // add title to 1st place in Main array
                    mainArr[0] = arrName[0];

                    // add first name to the 2nd plcae in the main array
                    mainArr[1] = arrName[1];

                    // add last name to the 5th place
                    mainArr[4] = arrName[2];

                    // add suffix to 6th place in main array
                    mainArr[5] = arrName[3];
                }


                else if (checkTitle(arrName[0]) == true & checkSuffix(arrName[3]) == false)
                {
                    // Dr. John Phillip Smith

                    // add title to the 1st place
                    mainArr[0] = arrName[0];

                    // add frist name to the 2nd place
                    mainArr[1] = arrName[1];

                    // add middle name to the 3rd place
                    mainArr[2] = arrName[2];

                    // add last name to the 5th place
                    mainArr[4] = arrName[3];

                    // add abbreviation to the 7th place
                    mainArr[6] = Strings.Left(mainArr[1], 1) + "." + Strings.Left(mainArr[2], 1) + ". " + mainArr[4];
                }


                else if (checkTitle(arrName[0]) == false & checkSuffix(arrName[3]) == true)
                {
                    // John Phillip Smith PhD

                    // add first name to the 2nd place
                    mainArr[1] = arrName[0];

                    // add middle name to the 3rd place
                    mainArr[2] = arrName[1];

                    // add last name to the 5th place
                    mainArr[4] = arrName[2];

                    // add suffix to the 6th place
                    mainArr[5] = arrName[3];

                    // add abbreviation to the 7th place
                    mainArr[6] = Strings.Left(mainArr[1], 1) + "." + Strings.Left(mainArr[2], 1) + ". " + mainArr[4];
                }

                else if (checkTitle(arrName[0]) == false & checkSuffix(arrName[3]) == false)
                {
                    // John Phillip Van Smith

                    // add first name to the 2nd place
                    mainArr[1] = arrName[0];

                    // add middle name to the 3rd place
                    mainArr[2] = arrName[1];

                    // add last name prefix in 4th place
                    mainArr[3] = arrName[2];

                    // add last name to the 5th plcae
                    mainArr[4] = arrName[3];

                    // add abbreviation to the 7th place
                    mainArr[6] = Strings.Left(mainArr[1], 1) + "." + Strings.Left(mainArr[2], 1) + ". " + mainArr[3] + " " + mainArr[4];


                }
            }

            else if (Information.UBound(arrName) == 4)
            {

                if (checkTitle(arrName[0]) == true & checkSuffix(arrName[4]) == true)
                {
                    // Dr. John Phillip Smith PhD

                    // add title to 1st place in Main array
                    mainArr[0] = arrName[0];

                    // add first name to the 2nd plcae in the main array
                    mainArr[1] = arrName[1];

                    // add middle name to the 3rd place
                    mainArr[2] = arrName[2];


                    // add last name to the 5th place
                    mainArr[4] = arrName[3];

                    // add suffix to 6th place in main array
                    mainArr[5] = arrName[4];

                    // add abbreviation to the 7th place
                    mainArr[6] = Strings.Left(mainArr[1], 1) + "." + Strings.Left(mainArr[2], 1) + ". " + mainArr[4];
                }


                else if (checkTitle(arrName[0]) == true & checkSuffix(arrName[4]) == false)
                {
                    // Dr. John Phillip Van Smith

                    // add title to the 1st place
                    mainArr[0] = arrName[0];

                    // add frist name to the 2nd place
                    mainArr[1] = arrName[1];

                    // add middle name to the 3rd place
                    mainArr[2] = arrName[2];

                    // add last name prefix in 4th place
                    mainArr[3] = arrName[3];

                    // add last name to the 5th place
                    mainArr[4] = arrName[4];

                    // add abbreviation to the 7th place
                    mainArr[6] = Strings.Left(mainArr[1], 1) + "." + Strings.Left(mainArr[2], 1) + ". " + mainArr[3] + " " + mainArr[4];
                }


                else if (checkTitle(arrName[0]) == false & checkSuffix(arrName[4]) == true)
                {
                    // John Phillip Van Smith PhD

                    // add first name to the 2nd place
                    mainArr[1] = arrName[0];

                    // add middle name to the 3rd place
                    mainArr[2] = arrName[1];

                    // add last name prefix in 4th place
                    mainArr[3] = arrName[2];

                    // add last name to the 5th place
                    mainArr[4] = arrName[3];

                    // add suffix to the 6th place
                    mainArr[5] = arrName[4];

                    // add abbreviation to the 7th place
                    mainArr[6] = Strings.Left(mainArr[1], 1) + "." + Strings.Left(mainArr[2], 1) + ". " + mainArr[3] + " " + mainArr[4];
                }

                else if (checkTitle(arrName[0]) == false & checkSuffix(arrName[4]) == false)
                {
                    // John Phillip Van Der Smith

                    // add first name to the 2nd place
                    mainArr[1] = arrName[0];

                    // add middle name to the 3rd place
                    mainArr[2] = arrName[1];

                    // add last name prefix in 4th place
                    mainArr[3] = arrName[2] + " " + arrName[3];

                    // add last name to the 5th plcae
                    mainArr[4] = arrName[4];

                    // add abbreviation to the 7th place
                    mainArr[6] = Strings.Left(mainArr[1], 1) + "." + Strings.Left(mainArr[2], 1) + ". " + mainArr[3] + " " + mainArr[4];


                }
            }

            else if (Information.UBound(arrName) >= 5)
            {


                if (checkTitle(arrName[0]) == true & checkSuffix(arrName[Information.UBound(arrName)]) == true)
                {
                    // Dr. John Phillip Van ... Smith PhD

                    // add title to 1st place in Main array
                    mainArr[0] = arrName[0];

                    // add first name to the 2nd plcae in the main array
                    mainArr[1] = arrName[1];

                    // add middle name to the 3rd place
                    mainArr[2] = arrName[2];

                    // add last name prefix in 4th place
                    for (int k = 3, loopTo = Information.UBound(arrName) - 2; k <= loopTo; k++)
                        mainArr[3] = mainArr[3] + " " + arrName[k];
                    // remove any extra leading and trailing spaces
                    mainArr[3] = Strings.Trim(mainArr[3]);


                    // add last name to the 5th place
                    mainArr[4] = arrName[Information.UBound(arrName) - 1];

                    // mainArr(5) = arrName(UBound(arrName) - 1)


                    // add abbreviation to the 7th place
                    mainArr[6] = Strings.Left(mainArr[1], 1) + "." + Strings.Left(mainArr[2], 1) + ". " + mainArr[3] + " " + mainArr[4];


                    // add suffix to 6th place in main array
                    mainArr[5] = arrName[Information.UBound(arrName)];
                }


                else if (checkTitle(arrName[0]) == true & checkSuffix(arrName[Information.UBound(arrName)]) == false)
                {
                    // Dr. John Phillip Van Der ... Smith

                    // add title to the 1st place
                    mainArr[0] = arrName[0];

                    // add first name to the 2nd place
                    mainArr[1] = arrName[1];

                    // add middle name to the 3rd place
                    mainArr[2] = arrName[2];

                    // add last name prefix in 4th place
                    for (int k = 3, loopTo1 = Information.UBound(arrName) - 1; k <= loopTo1; k++)
                        mainArr[3] = mainArr[3] + " " + arrName[k];
                    // remove any extra leading and trailing spaces
                    mainArr[3] = Strings.Trim(mainArr[3]);

                    // add last name to the 5th place
                    mainArr[4] = arrName[Information.UBound(arrName)];

                    // add abbreviation to the 7th place
                    mainArr[6] = Strings.Left(mainArr[1], 1) + "." + Strings.Left(mainArr[2], 1) + ". " + mainArr[3] + " " + mainArr[4];
                }


                else if (checkTitle(arrName[0]) == false & checkSuffix(arrName[Information.UBound(arrName)]) == true)
                {
                    // John Phillip Van Der ... Smith PhD

                    // add first name to the 2nd place
                    mainArr[1] = arrName[0];

                    // add middle name to the 3rd place
                    mainArr[2] = arrName[1];

                    // add last name prefix in 4th place
                    for (int k = 2, loopTo2 = Information.UBound(arrName) - 2; k <= loopTo2; k++)
                        mainArr[3] = mainArr[3] + " " + arrName[k];
                    // remove any extra leading and trailing spaces
                    mainArr[3] = Strings.Trim(mainArr[3]);

                    // add last name to the 5th place
                    mainArr[4] = arrName[Information.UBound(arrName) - 1];

                    // add suffix to the 6th place
                    mainArr[5] = arrName[Information.UBound(arrName)];

                    // add abbreviation to the 7th place
                    mainArr[6] = Strings.Left(mainArr[1], 1) + "." + Strings.Left(mainArr[2], 1) + ". " + mainArr[3] + " " + mainArr[4];
                }

                else if (checkTitle(arrName[0]) == false & checkSuffix(arrName[Information.UBound(arrName)]) == false)
                {
                    // John Phillip Van Der James ... Smith 

                    // add first name to the 2nd place
                    mainArr[1] = arrName[0];

                    // add middle name to the 3rd place
                    mainArr[2] = arrName[1];

                    // add last name prefix in 4th place
                    for (int k = 2, loopTo3 = Information.UBound(arrName) - 1; k <= loopTo3; k++)
                        mainArr[3] = mainArr[3] + " " + arrName[k];
                    // remove any extra leading and trailing spaces
                    mainArr[3] = Strings.Trim(mainArr[3]);

                    // add last name to the 5th plcae
                    mainArr[4] = arrName[Information.UBound(arrName)];

                    // add abbreviation to the 7th place
                    mainArr[6] = Strings.Left(mainArr[1], 1) + "." + Strings.Left(mainArr[2], 1) + ". " + mainArr[3] + " " + mainArr[4];

                }





            }



        }

        private void display()
        {

            try
            {


                CustomPanel2.Controls.Clear();



                Range displayRng = (Range)worksheet.Cells[1, 1];
                string[] arrRng;


                var outputColumn = default(int);

                arrRng = Strings.Split(txtSourceRange.Text, ",");

                var headerIndex = default(int);
                for (int i = 0, loopTo = Information.UBound(arrRng); i <= loopTo; i++)
                {

                    sourceRange = worksheet.get_Range(arrRng[i]);

                    for (int j = 1, loopTo1 = sourceRange.Rows.Count; j <= loopTo1; j++)
                    {

                        mainArr = new[] { "", "", "", "", "", "", "", "" };
                        Name = Conversions.ToString(sourceRange.Cells[j, 1].value);

                        nameSplitter();
                        string headerStr = "";
                        foreach (Control ctrl in CustomGroupBox7.Controls)
                        {
                            if (ctrl is System.Windows.Forms.CheckBox)
                            {
                                System.Windows.Forms.CheckBox chk = (System.Windows.Forms.CheckBox)ctrl;
                                if (chk.Checked)
                                {
                                    headerStr = headerStr + "," + chk.Text;
                                }
                            }
                        }

                        headerStr = headerStr.Replace("Select All,", string.Empty);
                        headerStr = Strings.Right(headerStr, Strings.Len(headerStr) - 1);

                        string[] arrHeaderStr = Strings.Split(headerStr, ",");
                        outputColumn = Information.UBound(arrHeaderStr) + 1;
                        for (int k = 0, loopTo2 = Information.UBound(arrHeaderStr); k <= loopTo2; k++)
                        {
                            worksheet.Cells[100000, k + 1].value = arrHeaderStr[k];

                            switch (arrHeaderStr[k] ?? "")
                            {
                                case var @case when @case == "Title":
                                    {
                                        headerIndex = 0;
                                        break;
                                    }
                                case var case1 when case1 == "First Name":
                                    {
                                        headerIndex = 1;
                                        break;
                                    }
                                case var case2 when case2 == "Middle Name":
                                    {
                                        headerIndex = 2;
                                        break;
                                    }
                                case var case3 when case3 == "Last Name Prefix":
                                    {
                                        headerIndex = 3;
                                        break;
                                    }
                                case var case4 when case4 == "Last Name":
                                    {
                                        headerIndex = 4;
                                        break;
                                    }
                                case var case5 when case5 == "Name Suffix":
                                    {
                                        headerIndex = 5;
                                        break;
                                    }
                                case var case6 when case6 == "Name Abbreviations":
                                    {
                                        headerIndex = 6;
                                        break;
                                    }

                            }

                            worksheet.Cells[100000 + j, k + 1].value = mainArr[headerIndex];

                        }





                    }

                    displayRng = worksheet.get_Range(worksheet.Cells[100000, 1], worksheet.Cells[100000 + sourceRange.Rows.Count, outputColumn]);

                    if (CB_Add_Header.Checked == false)
                    {

                        displayRng = worksheet.get_Range(worksheet.Cells[100001, 1], worksheet.Cells[100000 + sourceRange.Rows.Count, outputColumn]);

                    }

                }


                if (string.IsNullOrEmpty(txtSourceRange.Text) | displayRng is null)
                {
                    CustomPanel2.Controls.Clear();
                    return;
                }


                if (displayRng.Rows.Count > 50)
                {
                    displayRng = (Range)displayRng.Rows["1:50"];
                }


                double height;
                double width;

                if (displayRng.Rows.Count <= 4)
                {
                    height = CustomPanel2.Height / (double)displayRng.Rows.Count;
                }
                else
                {
                    height = 119d / 4d;
                }

                if (displayRng.Columns.Count <= 3)
                {
                    width = CustomPanel2.Width / (double)displayRng.Columns.Count;
                }
                else
                {
                    width = 260d / 3d;
                }






                for (int i = 1, loopTo3 = displayRng.Rows.Count; i <= loopTo3; i++)
                {
                    for (int j = 1, loopTo4 = displayRng.Columns.Count; j <= loopTo4; j++)
                    {
                        var label = new System.Windows.Forms.Label();
                        label.Text = Conversions.ToString(displayRng.Cells[i, j].Value);
                        label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * width), (int)Math.Round((i - 1) * height));
                        label.Height = (int)Math.Round(height);
                        label.Width = (int)Math.Round(width);
                        label.BorderStyle = BorderStyle.FixedSingle;
                        label.TextAlign = ContentAlignment.MiddleCenter;

                        if (CB_Keep_Formatting.Checked == true)
                        {
                            if (CB_Add_Header.Checked == true)
                            {

                                var cellColor = ColorTranslator.FromOle(Conversions.ToInteger(sourceRange.Cells[i - 1, 1].Font.Color));
                                var fillColor = ColorTranslator.FromOle(Conversions.ToInteger(sourceRange.Cells[i - 1, 1].interior.Color));


                                Range cell = (Range)sourceRange.Cells[i - 1, 1];

                                string cellFontName = Conversions.ToString(cell.Font.Name);
                                float cellFontSize = Convert.ToSingle(10);
                                var cellFontColor = ColorTranslator.FromOle(Conversions.ToInteger(cell.Font.Color));
                                var cellFontStyle = FontStyle.Regular;
                                if (Conversions.ToBoolean(cell.Font.Bold))
                                    cellFontStyle = cellFontStyle | FontStyle.Bold;
                                if (Conversions.ToBoolean(cell.Font.Italic))
                                    cellFontStyle = cellFontStyle | FontStyle.Italic;
                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(cell.Font.Underline, XlUnderlineStyle.xlUnderlineStyleNone, false)))
                                    cellFontStyle = cellFontStyle | FontStyle.Underline;

                                label.Font = new System.Drawing.Font(cellFontName, cellFontSize, cellFontStyle);


                                label.ForeColor = cellColor;
                                label.BackColor = fillColor;

                                // bold header
                                if (i == 1)
                                {

                                    cellFontStyle = cellFontStyle | FontStyle.Bold;
                                    label.Font = new System.Drawing.Font(cellFontName, cellFontSize, cellFontStyle);
                                }
                            }

                            else
                            {

                                var cellColor = ColorTranslator.FromOle(Conversions.ToInteger(sourceRange.Cells[i, 1].Font.Color));
                                var fillColor = ColorTranslator.FromOle(Conversions.ToInteger(sourceRange.Cells[i, 1].interior.Color));


                                Range cell = (Range)sourceRange.Cells[i, 1];

                                string cellFontName = Conversions.ToString(cell.Font.Name);
                                float cellFontSize = Convert.ToSingle(10);
                                var cellFontColor = ColorTranslator.FromOle(Conversions.ToInteger(cell.Font.Color));
                                var cellFontStyle = FontStyle.Regular;
                                if (Conversions.ToBoolean(cell.Font.Bold))
                                    cellFontStyle = cellFontStyle | FontStyle.Bold;
                                if (Conversions.ToBoolean(cell.Font.Italic))
                                    cellFontStyle = cellFontStyle | FontStyle.Italic;
                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(cell.Font.Underline, XlUnderlineStyle.xlUnderlineStyleNone, false)))
                                    cellFontStyle = cellFontStyle | FontStyle.Underline;

                                label.Font = new System.Drawing.Font(cellFontName, cellFontSize, cellFontStyle);


                                label.ForeColor = cellColor;
                                label.BackColor = fillColor;


                            }
                        }


                        else
                        {
                            label.BackColor = Color.Transparent;
                            label.ForeColor = default;



                        }


                        CustomPanel2.Controls.Add(label);
                    }
                }

                CustomPanel2.AutoScroll = true;

                worksheet.get_Range(displayRng.Cells[1, 1].offset((object)-1, (object)0), displayRng.Cells[displayRng.Rows.Count, displayRng.Columns.Count]).EntireRow.Delete();
            }

            catch (Exception ex)
            {

            }



        }



        private void uncheck_CB_Select_All()
        {
            foreach (Control ctrl in CustomGroupBox7.Controls)
            {
                if (ctrl is System.Windows.Forms.CheckBox)
                {
                    System.Windows.Forms.CheckBox chk = (System.Windows.Forms.CheckBox)ctrl;
                    // Do something with chk
                    if (chk.Checked == false)
                    {
                        CB_Select_All.Checked = false;
                    }
                }
            }
            display();

        }


        private void CB_Select_All_CheckedChanged(object sender, EventArgs e)
        {
            if (CB_Select_All.Checked == true)
            {
                foreach (Control ctrl in CustomGroupBox7.Controls)
                {
                    if (ctrl is System.Windows.Forms.CheckBox)
                    {
                        System.Windows.Forms.CheckBox chk = (System.Windows.Forms.CheckBox)ctrl;

                        chk.Checked = true;
                        chk.Enabled = false;

                    }
                }
                CB_Select_All.Checked = true;
                CB_Select_All.Enabled = true;
            }

            else if (CB_Select_All.Checked == false)
            {
                foreach (Control ctrl in CustomGroupBox7.Controls)
                {
                    if (ctrl is System.Windows.Forms.CheckBox)
                    {
                        System.Windows.Forms.CheckBox chk = (System.Windows.Forms.CheckBox)ctrl;

                        chk.Checked = false;
                        chk.Enabled = true;

                    }
                }


            }

            display();


        }

        private void CB_Title_CheckedChanged(object sender, EventArgs e)
        {

            uncheck_CB_Select_All();

        }

        private void CB_First_Name_CheckedChanged(object sender, EventArgs e)
        {
            uncheck_CB_Select_All();
        }

        private void CB_Middle_Name_CheckedChanged(object sender, EventArgs e)
        {
            uncheck_CB_Select_All();
        }

        private void CB_Last_Name_Prefix_CheckedChanged(object sender, EventArgs e)
        {
            uncheck_CB_Select_All();
        }

        private void CB_Last_Name_CheckedChanged(object sender, EventArgs e)
        {
            uncheck_CB_Select_All();
        }



        private void CB_Name_Abbreviations_CheckedChanged(object sender, EventArgs e)
        {
            uncheck_CB_Select_All();
        }



        private void CB_Name_Suffix_CheckedChanged(object sender, EventArgs e)
        {
            uncheck_CB_Select_All();
        }

        private void RB_Same_As_Source_Range_CheckedChanged(object sender, EventArgs e)
        {

            if (RB_Same_As_Source_Range.Checked == true)
            {

                txtDestRange.Enabled = false;
                destinationSelection.Enabled = false;
                lbl_destRange_Selection.Enabled = false;
            }

            else if (RB_Same_As_Source_Range.Checked == false)
            {

                txtDestRange.Enabled = true;
                destinationSelection.Enabled = true;
                lbl_destRange_Selection.Enabled = true;

            }

        }

        private void CB_Add_Header_CheckedChanged(object sender, EventArgs e)
        {
            display();
        }

        private void CB_Keep_Formatting_CheckedChanged(object sender, EventArgs e)
        {
            display();
        }

        private void Form17DivideNames_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void Form17DivideNames_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form17DivideNames_Shown(object sender, EventArgs e)
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