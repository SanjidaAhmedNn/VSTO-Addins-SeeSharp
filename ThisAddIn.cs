using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace VSTO_Addins
{

    public partial class ThisAddIn
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
        public static Excel.Worksheet worksheet;
        // Class-level variable for your form
        private Form36 Form = null;
        private Form38 Form2 = null;
        private Form40 Form3 = null;
        public Range src_rng;
        public Range src_rng1;
        public Range src_rng2;
        public Range src_rng3;
        public Range src_rng4;
        public Range src_rng5;
        public Range des_rng;
        public Range des_rng1;
        public Range des_rng2;
        public Range des_rng3;
        public Range des_rng4;
        public Range des_rng5;

        public string sheetName3;
        public string sheetName4;

        public Range range1;
        public Range range2;

        private Excel.Worksheet _wsEvent1;

        private Excel.Worksheet wsEvent1
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _wsEvent1;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_wsEvent1 != null)
                {
                    _wsEvent1.SelectionChange -= wsEvent1_SelectionChange;
                }

                _wsEvent1 = value;
                if (_wsEvent1 != null)
                {
                    _wsEvent1.SelectionChange += wsEvent1_SelectionChange;
                }
            }
        }
        private Excel.Worksheet _wsEvent2;

        private Excel.Worksheet wsEvent2
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _wsEvent2;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_wsEvent2 != null)
                {
                    _wsEvent2.SelectionChange -= wsEvent2_SelectionChange;
                }

                _wsEvent2 = value;
                if (_wsEvent2 != null)
                {
                    _wsEvent2.SelectionChange += wsEvent2_SelectionChange;
                }
            }
        }
        private Excel.Worksheet _wsEvent3;

        private Excel.Worksheet wsEvent3
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _wsEvent3;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_wsEvent3 != null)
                {
                    _wsEvent3.SelectionChange -= wsEvent3_SelectionChange;
                }

                _wsEvent3 = value;
                if (_wsEvent3 != null)
                {
                    _wsEvent3.SelectionChange += wsEvent3_SelectionChange;
                }
            }
        }
        private Excel.Worksheet _wsEvent4_1;

        private Excel.Worksheet wsEvent4_1
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _wsEvent4_1;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_wsEvent4_1 != null)
                {
                    _wsEvent4_1.Change -= wsEvent4_SelectionChange;
                }

                _wsEvent4_1 = value;
                if (_wsEvent4_1 != null)
                {
                    _wsEvent4_1.Change += wsEvent4_SelectionChange;
                }
            }
        }
        private Excel.Worksheet _wsEvent4_2;

        private Excel.Worksheet wsEvent4_2
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _wsEvent4_2;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_wsEvent4_2 != null)
                {
                    _wsEvent4_2.Change -= wsEvent4_2_SelectionChange;
                }

                _wsEvent4_2 = value;
                if (_wsEvent4_2 != null)
                {
                    _wsEvent4_2.Change += wsEvent4_2_SelectionChange;
                }
            }
        }
        private Excel.Worksheet _wsEvent4_3;

        private Excel.Worksheet wsEvent4_3
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _wsEvent4_3;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_wsEvent4_3 != null)
                {
                    _wsEvent4_3.Change -= wsEvent4_3_SelectionChange;
                }

                _wsEvent4_3 = value;
                if (_wsEvent4_3 != null)
                {
                    _wsEvent4_3.Change += wsEvent4_3_SelectionChange;
                }
            }
        }
        private Excel.Worksheet _wsEvent4_4;

        private Excel.Worksheet wsEvent4_4
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _wsEvent4_4;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_wsEvent4_4 != null)
                {
                    _wsEvent4_4.Change -= wsEvent4_4_SelectionChange;
                }

                _wsEvent4_4 = value;
                if (_wsEvent4_4 != null)
                {
                    _wsEvent4_4.Change += wsEvent4_4_SelectionChange;
                }
            }
        }
        private Excel.Worksheet _wsEvent4_5;

        private Excel.Worksheet wsEvent4_5
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _wsEvent4_5;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_wsEvent4_5 != null)
                {
                    _wsEvent4_5.Change -= wsEvent4_5_SelectionChange;
                }

                _wsEvent4_5 = value;
                if (_wsEvent4_5 != null)
                {
                    _wsEvent4_5.Change += wsEvent4_5_SelectionChange;
                }
            }
        }
        private Excel.Worksheet wsEvent5;


        private Excel.Worksheet _CurrentSheet;

        private Excel.Worksheet CurrentSheet
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _CurrentSheet;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _CurrentSheet = value;
            }
        }
        private Excel.Workbook _WorkbookEvents;

        private Excel.Workbook WorkbookEvents
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _WorkbookEvents;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_WorkbookEvents != null)
                {
                    _WorkbookEvents.SheetActivate -= WorkbookEvents_SheetActivate;
                }

                _WorkbookEvents = value;
                if (_WorkbookEvents != null)
                {
                    _WorkbookEvents.SheetActivate += WorkbookEvents_SheetActivate;
                }
            }
        }



        private void ThisAddIn_Startup(object sender, EventArgs e)
        {

            Globals.ThisAddIn.Application.DisplayAlerts = false;
            Application.EnableEvents = true;
            GlobalModule.form_flag = false;
            GlobalModule.sessionflag1 = true;
            GlobalModule.sessionflag2 = true;


            Globals.ThisAddIn.Application.WorkbookActivate += Workbook_Activated;
        }

        private void Workbook_Activated(Excel.Workbook Wb)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workBook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workBook.ActiveSheet;
            RemovePreviousEventHandler(); // We'll define this function to ensure we don't attach multiple handlers.

            CheckForNewwWorksheet();
            HideNewwwWorksheet();

            if (GlobalModule.Flag1 == true)
            {

                worksheet.get_Range("B1").Select();  // Randomly select a cell. If nothing is selected, addhandler show error

                // Define an array of type Excel.Worksheet
                Excel.Worksheet[] sheetsArray;

                // Resize the array based on the number of sheets
                sheetsArray = new Excel.Worksheet[workBook.Worksheets.Count + 1];

                if (GlobalModule.SR1.Contains("Active Workbook"))
                {
                    // Adding sheet_selectionchange event
                    worksheet.SelectionChange += sheet_SelectionChange1;

                    WorkbookEvents = Globals.ThisAddIn.Application.ActiveWorkbook;
                    CurrentSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                }


                else if (GlobalModule.SR1 == "Select Range" | GlobalModule.SR1.Contains("Active Sheet"))
                {
                    for (int i = 1, loopTo1 = excelApp.ActiveWorkbook.Worksheets.Count; i <= loopTo1; i++)
                    {
                        sheetsArray[i] = (Excel.Worksheet)workBook.Worksheets[i];
                        if ((GlobalModule.shName1 ?? "") == (sheetsArray[i].Name ?? ""))
                        {
                            wsEvent1 = sheetsArray[i];
                            // AddHandler sheetsArray(i).SelectionChange, AddressOf sheet_SelectionChange1
                            // src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                        }

                    }
                }

                else
                {

                    for (int i = 1, loopTo = excelApp.ActiveWorkbook.Worksheets.Count; i <= loopTo; i++)
                    {
                        sheetsArray[i] = (Excel.Worksheet)workBook.Worksheets[i];
                        if ((GlobalModule.SR1 ?? "") == (sheetsArray[i].Name ?? ""))
                        {
                            sheetsArray[i].SelectionChange += sheet_SelectionChange1;
                            // src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                        }

                    }
                    GlobalModule.GB_CB_Source1 = worksheet.get_Range("A1", worksheet.Cells[excelApp.Rows.Count, excelApp.Columns.Count]).get_Address();

                }
            }

            if (GlobalModule.Flag2 == true)
            {

                worksheet.get_Range("B1").Select();  // Randomly select a cell. If nothing is selected, addhandler show error

                // Define an array of type Excel.Worksheet
                Excel.Worksheet[] sheetsArray;

                // Resize the array based on the number of sheets
                sheetsArray = new Excel.Worksheet[workBook.Worksheets.Count + 1];

                if (GlobalModule.SR2.Contains("Active Workbook"))
                {
                    // Add sheetselectionchange_event
                    // Assuming you're working with the active workbook:
                    worksheet.SelectionChange += sheet_SelectionChange2;
                    WorkbookEvents = Globals.ThisAddIn.Application.ActiveWorkbook;
                    CurrentSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                }

                // For i = 1 To excelApp.ActiveWorkbook.Worksheets.Count
                // MsgBox(excelApp.ActiveWorkbook.Worksheets.Count)
                // sheetsArray(i) = CType(workBook.Worksheets(i), Excel.Worksheet)

                // 'Needs to add handler to each worksheet
                // If sheetsArray(i).Visible = Excel.XlSheetVisibility.xlSheetVisible Then
                // AddHandler sheetsArray(i).SelectionChange, AddressOf sheet_SelectionChange
                // End If
                // Next


                else if (GlobalModule.SR2 == "Select Range" | GlobalModule.SR2.Contains("Active Sheet"))
                {

                    for (int i = 1, loopTo3 = excelApp.ActiveWorkbook.Worksheets.Count; i <= loopTo3; i++)
                    {
                        sheetsArray[i] = (Excel.Worksheet)workBook.Worksheets[i];
                        if ((GlobalModule.shName2 ?? "") == (sheetsArray[i].Name ?? ""))
                        {
                            wsEvent2 = sheetsArray[i];
                            // AddHandler sheetsArray(i).SelectionChange, AddressOf sheet_SelectionChange2
                            // src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                        }

                    }
                }
                else
                {

                    for (int i = 1, loopTo2 = excelApp.ActiveWorkbook.Worksheets.Count; i <= loopTo2; i++)
                    {
                        sheetsArray[i] = (Excel.Worksheet)workBook.Worksheets[i];
                        if ((GlobalModule.SR2 ?? "") == (sheetsArray[i].Name ?? ""))
                        {
                            sheetsArray[i].SelectionChange += sheet_SelectionChange2;
                            // src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                        }
                        // i = i + 1

                    }
                    GlobalModule.GB_CB_Source2 = worksheet.get_Range("A1", worksheet.Cells[excelApp.Rows.Count, excelApp.Columns.Count]).get_Address();

                }
            }

            if (GlobalModule.Flag3 == true)
            {

                worksheet.get_Range("B1").Select();  // Randomly select a cell. If nothing is selected, addhandler show error

                // Define an array of type Excel.Worksheet
                Excel.Worksheet[] sheetsArray;

                // Resize the array based on the number of sheets
                sheetsArray = new Excel.Worksheet[workBook.Worksheets.Count + 1];

                if (GlobalModule.SR3.Contains("Active Workbook"))
                {

                    // Assuming you're working with the active workbook:
                    worksheet.SelectionChange += sheet_SelectionChange3;
                    WorkbookEvents = Globals.ThisAddIn.Application.ActiveWorkbook;
                    CurrentSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                }

                // For i = 1 To excelApp.ActiveWorkbook.Worksheets.Count

                // sheetsArray(i) = CType(workBook.Worksheets(i), Excel.Worksheet)

                // 'Needs to add handler to each worksheet
                // If sheetsArray(i).Visible = Excel.XlSheetVisibility.xlSheetVisible Then
                // AddHandler sheetsArray(i).SelectionChange, AddressOf sheet_SelectionChange
                // End If
                // Next


                else if (GlobalModule.SR3 == "Select Range" | GlobalModule.SR3.Contains("Active Sheet"))
                {

                    for (int i = 1, loopTo5 = excelApp.ActiveWorkbook.Worksheets.Count; i <= loopTo5; i++)
                    {
                        sheetsArray[i] = (Excel.Worksheet)workBook.Worksheets[i];
                        if ((GlobalModule.shName2 ?? "") == (sheetsArray[i].Name ?? ""))
                        {
                            wsEvent3 = sheetsArray[i];
                            // AddHandler sheetsArray(i).SelectionChange, AddressOf sheet_SelectionChange3
                            // src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                        }
                        // i = i + 1

                    }
                }
                else
                {

                    for (int i = 1, loopTo4 = excelApp.ActiveWorkbook.Worksheets.Count; i <= loopTo4; i++)
                    {
                        sheetsArray[i] = (Excel.Worksheet)workBook.Worksheets[i];
                        if ((GlobalModule.SR3 ?? "") == (sheetsArray[i].Name ?? ""))
                        {
                            sheetsArray[i].SelectionChange += sheet_SelectionChange3;
                            // src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                        }
                        // i = i + 1

                    }
                    GlobalModule.GB_CB_Source3 = worksheet.get_Range("A1", worksheet.Cells[excelApp.Rows.Count, excelApp.Columns.Count]).get_Address();

                }

            }

            if (GlobalModule.Flag_CreateDDDL == true)
            {
                Excel.Worksheet ws1;
                Excel.Worksheet ws2;
                foreach (var ws in excelApp.ActiveWorkbook.Worksheets)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.name, "MySpecialSheet", false)))
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(ws.Range("A1").Value, "", false)))
                        {
                            GlobalModule.Variable1 = ws.Range("A1").Value.ToString();
                            GlobalModule.Variable2 = ws.Range("A2").Value.ToString();
                            GlobalModule.Header = Conversions.ToBoolean(ws.Range("A3").Value.ToString());
                            GlobalModule.Ascending = Conversions.ToBoolean(ws.Range("A4").Value.ToString());
                            GlobalModule.Descending = Conversions.ToBoolean(ws.Range("A5").Value.ToString());
                            GlobalModule.TextConvert = Conversions.ToBoolean(ws.Range("A6").Value.ToString());
                            GlobalModule.OptionType = Conversions.ToBoolean(ws.Range("A7").Value.ToString());
                            GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(ws.Range("A8").Value.ToString());
                            GlobalModule.Flag_CreateDDDL = Conversions.ToBoolean(ws.Range("A9").value.ToString());
                            sheetName3 = ws.Range("A10").value.ToString();
                            sheetName4 = ws.Range("A11").value.ToString();



                            ws1 = (Excel.Worksheet)workBook.Worksheets[sheetName4];
                            ws2 = (Excel.Worksheet)workBook.Worksheets[sheetName3];
                            src_rng1 = ws2.get_Range(GlobalModule.Variable1);
                            des_rng1 = ws1.get_Range(GlobalModule.Variable2);
                            // AddHandler ws1.Change, AddressOf worksheet5_1_Change
                            wsEvent4_1 = ws1;

                        }

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(ws.Range("B1").Value, "", false)))
                        {
                            GlobalModule.Variable1 = ws.Range("B1").Value.ToString();
                            GlobalModule.Variable2 = ws.Range("B2").Value.ToString();
                            GlobalModule.Header = Conversions.ToBoolean(ws.Range("B3").Value.ToString());
                            GlobalModule.Ascending = Conversions.ToBoolean(ws.Range("B4").Value.ToString());
                            GlobalModule.Descending = Conversions.ToBoolean(ws.Range("B5").Value.ToString());
                            GlobalModule.TextConvert = Conversions.ToBoolean(ws.Range("B6").Value.ToString());
                            GlobalModule.OptionType = Conversions.ToBoolean(ws.Range("B7").Value.ToString());
                            GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(ws.Range("B8").Value.ToString());
                            GlobalModule.Flag_CreateDDDL = Conversions.ToBoolean(ws.Range("B9").value.ToString());
                            sheetName3 = ws.Range("B10").value.ToString();
                            sheetName4 = ws.Range("B11").value.ToString();
                            ws1 = (Excel.Worksheet)workBook.Worksheets[sheetName4];

                            ws2 = (Excel.Worksheet)workBook.Worksheets[sheetName3];

                            src_rng2 = ws2.get_Range(GlobalModule.Variable1);
                            des_rng2 = ws1.get_Range(GlobalModule.Variable2);

                            wsEvent4_2 = ws1;


                        }

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(ws.Range("C1").Value, "", false)))
                        {
                            GlobalModule.Variable1 = ws.Range("C1").Value.ToString();
                            GlobalModule.Variable2 = ws.Range("C2").Value.ToString();
                            GlobalModule.Header = Conversions.ToBoolean(ws.Range("C3").Value.ToString());
                            GlobalModule.Ascending = Conversions.ToBoolean(ws.Range("C4").Value.ToString());
                            GlobalModule.Descending = Conversions.ToBoolean(ws.Range("C5").Value.ToString());
                            GlobalModule.TextConvert = Conversions.ToBoolean(ws.Range("C6").Value.ToString());
                            GlobalModule.OptionType = Conversions.ToBoolean(ws.Range("C7").Value.ToString());
                            GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(ws.Range("C8").Value.ToString());
                            GlobalModule.Flag_CreateDDDL = Conversions.ToBoolean(ws.Range("C9").value.ToString());
                            sheetName3 = ws.Range("C10").value.ToString();
                            sheetName4 = ws.Range("C11").value.ToString();
                            ws1 = (Excel.Worksheet)workBook.Worksheets[sheetName4];

                            ws2 = (Excel.Worksheet)workBook.Worksheets[sheetName3];
                            src_rng3 = ws2.get_Range(GlobalModule.Variable1);
                            des_rng3 = ws1.get_Range(GlobalModule.Variable2);

                            wsEvent4_3 = ws1;

                        }

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(ws.Range("D1").Value, "", false)))
                        {
                            GlobalModule.Variable1 = ws.Range("D1").Value.ToString();
                            GlobalModule.Variable2 = ws.Range("D2").Value.ToString();
                            GlobalModule.Header = Conversions.ToBoolean(ws.Range("D3").Value.ToString());
                            GlobalModule.Ascending = Conversions.ToBoolean(ws.Range("D4").Value.ToString());
                            GlobalModule.Descending = Conversions.ToBoolean(ws.Range("D5").Value.ToString());
                            GlobalModule.TextConvert = Conversions.ToBoolean(ws.Range("D6").Value.ToString());
                            GlobalModule.OptionType = Conversions.ToBoolean(ws.Range("D7").Value.ToString());
                            GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(ws.Range("D8").Value.ToString());
                            GlobalModule.Flag_CreateDDDL = Conversions.ToBoolean(ws.Range("D9").value.ToString());
                            sheetName3 = ws.Range("D10").value.ToString();
                            sheetName4 = ws.Range("D11").value.ToString();

                            ws1 = (Excel.Worksheet)workBook.Worksheets[sheetName4];
                            ws2 = (Excel.Worksheet)workBook.Worksheets[sheetName3];
                            src_rng4 = ws2.get_Range(GlobalModule.Variable1);
                            des_rng4 = ws1.get_Range(GlobalModule.Variable2);

                            wsEvent4_4 = ws1;

                        }

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(ws.Range("E1").Value, "", false)))
                        {
                            GlobalModule.Variable1 = ws.Range("E1").Value.ToString();
                            GlobalModule.Variable2 = ws.Range("E2").Value.ToString();
                            GlobalModule.Header = Conversions.ToBoolean(ws.Range("E3").Value.ToString());
                            GlobalModule.Ascending = Conversions.ToBoolean(ws.Range("E4").Value.ToString());
                            GlobalModule.Descending = Conversions.ToBoolean(ws.Range("E5").Value.ToString());
                            GlobalModule.TextConvert = Conversions.ToBoolean(ws.Range("E6").Value.ToString());
                            GlobalModule.OptionType = Conversions.ToBoolean(ws.Range("E7").Value.ToString());
                            GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(ws.Range("E8").Value.ToString());
                            GlobalModule.Flag_CreateDDDL = Conversions.ToBoolean(ws.Range("E9").value.ToString());
                            sheetName3 = ws.Range("E10").value.ToString();
                            sheetName4 = ws.Range("E11").value.ToString();

                            ws1 = (Excel.Worksheet)workBook.Worksheets[sheetName4];
                            ws2 = (Excel.Worksheet)workBook.Worksheets[sheetName3];
                            src_rng5 = ws2.get_Range(GlobalModule.Variable1);
                            des_rng5 = ws1.get_Range(GlobalModule.Variable2);

                            wsEvent4_5 = ws1;
                        }
                    }
                }


            }

            if (GlobalModule.Flag_Picture == true)
            {
                // Define an array of type Excel.Worksheet
                Excel.Worksheet[] sheetsArray;

                // Resize the array based on the number of sheets
                sheetsArray = new Excel.Worksheet[workBook.Worksheets.Count + 1];
                for (int i = 1, loopTo6 = excelApp.ActiveWorkbook.Worksheets.Count; i <= loopTo6; i++)
                {
                    sheetsArray[i] = (Excel.Worksheet)workBook.Worksheets[i];
                    if ((GlobalModule.sheetName2 ?? "") == (sheetsArray[i].Name ?? ""))
                    {
                        // wsEvent5 = DirectCast(sheetsArray(i), Excel.Worksheet)
                        excelApp.get_Range(GlobalModule.Des_Rng_of_PictureDDL).Columns[2].ColumnWidth = excelApp.get_Range(GlobalModule.Src_Rng_of_PictureDDL).Columns[2].ColumnWidth;
                        excelApp.get_Range(GlobalModule.Des_Rng_of_PictureDDL).Rows.RowHeight = excelApp.get_Range(GlobalModule.Src_Rng_of_PictureDDL).RowHeight;

                        // worksheet7_Change(Target)
                        // worksheet6_Change(Target)
                        // AddHandler sheetsArray(i).Change, AddressOf worksheet7_Change
                        sheetsArray[i].Change += worksheet6_Change;
                        // MsgBox(sheetsArray(i).Name)
                        // worksheet.Change
                        // src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                    }
                    // i = i + 1

                }
                // wsEvent4 = DirectCast(worksheet, Excel.Worksheet)
                // AddHandler worksheet.Change, AddressOf worksheet6_Change
            }



        }

        // This event will trigger when a cell in the worksheet is selected.
        private void wsEvent1_SelectionChange(Range Target)
        {
            // For testing purposes, we'll just show a message box.
            // MsgBox("Cell selected: " & Target.Address)
            sheet_SelectionChange1(Target);
        }

        private void wsEvent2_SelectionChange(Range Target)
        {
            // For testing purposes, we'll just show a message box.
            // MsgBox("Cell selected: " & Target.Address)
            sheet_SelectionChange2(Target);
        }

        private void wsEvent3_SelectionChange(Range Target)
        {
            // For testing purposes, we'll just show a message box.
            // MsgBox("Cell selected: " & Target.Address)
            sheet_SelectionChange3(Target);
        }

        // For Create Dynamic List
        private void wsEvent4_SelectionChange(Range Target)
        {
            try
            {
                excelApp = Globals.ThisAddIn.Application;
                var workBook = excelApp.ActiveWorkbook;
                // Dim worksheet As Excel.Worksheet = workBook.ActiveSheet
                bool i = false;

                var targetsheet = default(Excel.Worksheet);
                foreach (var ws in excelApp.ActiveWorkbook.Worksheets)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.name, "MySpecialSheet", false)))
                    {
                        targetsheet = (Excel.Worksheet)ws;
                        i = true;
                        break;
                    }
                }

                if (i == true)
                {
                    GlobalModule.Header = Conversions.ToBoolean(targetsheet.get_Range("A3").get_Value().ToString());
                    GlobalModule.Ascending = Conversions.ToBoolean(targetsheet.get_Range("A4").get_Value().ToString());
                    GlobalModule.Descending = Conversions.ToBoolean(targetsheet.get_Range("A5").get_Value().ToString());
                    GlobalModule.TextConvert = Conversions.ToBoolean(targetsheet.get_Range("A6").get_Value().ToString());
                    GlobalModule.OptionType = Conversions.ToBoolean(targetsheet.get_Range("A7").get_Value().ToString());
                    GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetsheet.get_Range("A8").get_Value().ToString());
                    GlobalModule.sheetName10 = targetsheet.get_Range("A10").get_Value().ToString();
                    GlobalModule.sheetName11 = targetsheet.get_Range("A11").get_Value().ToString();

                    src_rng = src_rng1;
                    des_rng = des_rng1;
                    // If excelApp.Intersect(Target, des_rng) IsNot Nothing Then
                    // MsgBox(1)
                    worksheet5_2_Change(Target);
                }
            }
            catch (Exception ex)
            {
            }

        }

        private void wsEvent4_2_SelectionChange(Range Target)
        {
            try
            {
                excelApp = Globals.ThisAddIn.Application;
                var workBook = excelApp.ActiveWorkbook;
                // Dim worksheet As Excel.Worksheet = workBook.ActiveSheet
                bool i = false;

                var targetsheet = default(Excel.Worksheet);
                foreach (var ws in excelApp.ActiveWorkbook.Worksheets)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.name, "MySpecialSheet", false)))
                    {
                        targetsheet = (Excel.Worksheet)ws;
                        i = true;
                        break;
                    }
                }
                if (i == true)
                {
                    GlobalModule.Header = Conversions.ToBoolean(targetsheet.get_Range("B3").get_Value().ToString());
                    GlobalModule.Ascending = Conversions.ToBoolean(targetsheet.get_Range("B4").get_Value().ToString());
                    GlobalModule.Descending = Conversions.ToBoolean(targetsheet.get_Range("B5").get_Value().ToString());
                    GlobalModule.TextConvert = Conversions.ToBoolean(targetsheet.get_Range("B6").get_Value().ToString());
                    GlobalModule.OptionType = Conversions.ToBoolean(targetsheet.get_Range("B7").get_Value().ToString());
                    GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetsheet.get_Range("B8").get_Value().ToString());
                    GlobalModule.sheetName10 = targetsheet.get_Range("B10").get_Value().ToString();
                    GlobalModule.sheetName11 = targetsheet.get_Range("B11").get_Value().ToString();

                    src_rng = src_rng2;
                    des_rng = des_rng2;

                    // If excelApp.Intersect(Target, des_rng) IsNot Nothing Then
                    worksheet5_2_Change(Target);
                }
            }
            catch (Exception ex)
            {
            }

        }

        private void wsEvent4_3_SelectionChange(Range Target)
        {
            try
            {
                excelApp = Globals.ThisAddIn.Application;
                var workBook = excelApp.ActiveWorkbook;
                // Dim worksheet As Excel.Worksheet = workBook.ActiveSheet
                bool i = false;

                var targetsheet = default(Excel.Worksheet);
                foreach (var ws in excelApp.ActiveWorkbook.Worksheets)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.name, "MySpecialSheet", false)))
                    {
                        targetsheet = (Excel.Worksheet)ws;
                        i = true;
                        break;
                    }
                }
                if (i == true)
                {
                    GlobalModule.Header = Conversions.ToBoolean(targetsheet.get_Range("C3").get_Value().ToString());
                    GlobalModule.Ascending = Conversions.ToBoolean(targetsheet.get_Range("C4").get_Value().ToString());
                    GlobalModule.Descending = Conversions.ToBoolean(targetsheet.get_Range("C5").get_Value().ToString());
                    GlobalModule.TextConvert = Conversions.ToBoolean(targetsheet.get_Range("C6").get_Value().ToString());
                    GlobalModule.OptionType = Conversions.ToBoolean(targetsheet.get_Range("C7").get_Value().ToString());
                    GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetsheet.get_Range("C8").get_Value().ToString());
                    GlobalModule.sheetName10 = targetsheet.get_Range("C10").get_Value().ToString();
                    GlobalModule.sheetName11 = targetsheet.get_Range("C11").get_Value().ToString();

                    src_rng = src_rng3;
                    des_rng = des_rng3;

                    // If excelApp.Intersect(Target, des_rng) IsNot Nothing Then
                    worksheet5_2_Change(Target);
                }
            }
            catch (Exception ex)
            {
            }

        }

        private void wsEvent4_4_SelectionChange(Range Target)
        {
            try
            {
                excelApp = Globals.ThisAddIn.Application;
                var workBook = excelApp.ActiveWorkbook;
                // Dim worksheet As Excel.Worksheet = workBook.ActiveSheet
                bool i = false;

                var targetsheet = default(Excel.Worksheet);
                foreach (var ws in excelApp.ActiveWorkbook.Worksheets)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.name, "MySpecialSheet", false)))
                    {
                        targetsheet = (Excel.Worksheet)ws;
                        i = true;
                        break;
                    }
                }

                if (i == true)
                {
                    GlobalModule.Header = Conversions.ToBoolean(targetsheet.get_Range("D3").get_Value().ToString());
                    GlobalModule.Ascending = Conversions.ToBoolean(targetsheet.get_Range("D4").get_Value().ToString());
                    GlobalModule.Descending = Conversions.ToBoolean(targetsheet.get_Range("D5").get_Value().ToString());
                    GlobalModule.TextConvert = Conversions.ToBoolean(targetsheet.get_Range("D6").get_Value().ToString());
                    GlobalModule.OptionType = Conversions.ToBoolean(targetsheet.get_Range("D7").get_Value().ToString());
                    GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetsheet.get_Range("D8").get_Value().ToString());
                    GlobalModule.sheetName10 = targetsheet.get_Range("D10").get_Value().ToString();
                    GlobalModule.sheetName11 = targetsheet.get_Range("D11").get_Value().ToString();

                    src_rng = src_rng4;
                    des_rng = des_rng4;


                    // If excelApp.Intersect(Target, des_rng) IsNot Nothing Then
                    worksheet5_2_Change(Target);
                }
            }
            catch (Exception ex)
            {
            }

        }

        private void wsEvent4_5_SelectionChange(Range Target)
        {
            try
            {
                excelApp = Globals.ThisAddIn.Application;
                var workBook = excelApp.ActiveWorkbook;
                // Dim worksheet As Excel.Worksheet = workBook.ActiveSheet
                bool i = false;

                var targetsheet = default(Excel.Worksheet);
                foreach (var ws in excelApp.ActiveWorkbook.Worksheets)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.name, "MySpecialSheet", false)))
                    {
                        targetsheet = (Excel.Worksheet)ws;
                        i = true;
                        break;
                    }
                }
                if (i == true)
                {
                    GlobalModule.Header = Conversions.ToBoolean(targetsheet.get_Range("E3").get_Value().ToString());
                    GlobalModule.Ascending = Conversions.ToBoolean(targetsheet.get_Range("E4").get_Value().ToString());
                    GlobalModule.Descending = Conversions.ToBoolean(targetsheet.get_Range("E5").get_Value().ToString());
                    GlobalModule.TextConvert = Conversions.ToBoolean(targetsheet.get_Range("E6").get_Value().ToString());
                    GlobalModule.OptionType = Conversions.ToBoolean(targetsheet.get_Range("E7").get_Value().ToString());
                    GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetsheet.get_Range("E8").get_Value().ToString());
                    GlobalModule.sheetName10 = targetsheet.get_Range("E10").get_Value().ToString();
                    GlobalModule.sheetName11 = targetsheet.get_Range("E11").get_Value().ToString();

                    src_rng = src_rng5;
                    des_rng = des_rng5;


                    // If excelApp.Intersect(Target, des_rng) IsNot Nothing Then
                    worksheet5_2_Change(Target);
                }
            }
            catch (Exception ex)
            {
            }
        }


        // For Picture Drop-down List
        private void wsEvent5_SelectionChange(Range Target)
        {

            excelApp.get_Range(GlobalModule.Des_Rng_of_PictureDDL).Columns[2].ColumnWidth = excelApp.get_Range(GlobalModule.Src_Rng_of_PictureDDL).Columns[2].ColumnWidth;
            excelApp.get_Range(GlobalModule.Des_Rng_of_PictureDDL).Rows.RowHeight = excelApp.get_Range(GlobalModule.Src_Rng_of_PictureDDL).RowHeight;

            // worksheet7_Change(Target)
            worksheet6_Change(Target);

        }


        // Event handler for when any sheet in the workbook is activated
        private void WorkbookEvents_SheetActivate(object Sh)
        {
            // Detach event from previous sheet
            if (CurrentSheet is not null)
            {

                CurrentSheet.SelectionChange -= sheet_SelectionChange1;
            }

            // Attach event to the new active sheet
            CurrentSheet = (Excel.Worksheet)Sh;
            CurrentSheet.SelectionChange += sheet_SelectionChange1;

        }

        private void RemovePreviousEventHandler()
        {

            // This function ensures that we remove previously attached event handlers
            // to avoid multiple event triggers. 
            // This is a simplistic approach and may need refining based on your exact needs.
            foreach (var wb in excelApp.Workbooks)
            {
                foreach (Excel.Worksheet currentWorksheet in (System.Collections.IEnumerable)wb.Worksheets)
                {
                    worksheet = currentWorksheet;
                    worksheet.SelectionChange -= sheet_SelectionChange1;
                    worksheet.SelectionChange -= sheet_SelectionChange2;
                    worksheet.SelectionChange -= sheet_SelectionChange3;
                }
            }
        }

        private void CheckForNewwWorksheet()
        {

            excelApp = Globals.ThisAddIn.Application;
            var workBook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workBook.ActiveSheet;
            // excelApp = Globals.ThisAddIn.Application

            // Loop through each worksheet in the active workbook
            foreach (var ws in excelApp.ActiveWorkbook.Worksheets)
            {
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.Name, "Newwwwwwwwww", false)))
                {
                    // If worksheet "Neww" is found, store the value of A1 in the variable

                    GlobalModule.GB_CB_Source1 = ws.Range("A2").Value.ToString();
                    GlobalModule.SR1 = ws.Range("A3").Value.ToString();
                    GlobalModule.Horizontal1 = Conversions.ToBoolean(ws.Range("A4").Value.ToString());
                    GlobalModule.Separator1 = ws.Range("A5").Value.ToString();
                    GlobalModule.Search1 = Conversions.ToBoolean(ws.Range("A6").Value.ToString());
                    GlobalModule.Flag1 = Conversions.ToBoolean(ws.Range("A7").Value.ToString());
                    GlobalModule.shName1 = ws.Range("A9").Value.ToString();
                    GlobalModule.RangeType1 = ws.Range("B2").value.ToString();
                    // TargetVar = ws.Range("A8").Value.ToString

                    var src_rng = worksheet.get_Range(GlobalModule.GB_CB_Source1);
                    GlobalModule.TType1 = "";
                    // Exit Sub ' No need to check other sheets once "Neww" is found

                }
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.Name, "SoftekoSofteko", false)))   // For checkbox drop-down list
                {


                    // If worksheet "Neww" is found, store the value of A1 in the variable
                    GlobalModule.GB_CB_Source2 = ws.Range("A2").Value.ToString();

                    GlobalModule.SR2 = ws.Range("A3").Value.ToString();
                    GlobalModule.Horizontal2 = Conversions.ToBoolean(ws.Range("A4").Value.ToString());
                    GlobalModule.Separator2 = ws.Range("A5").Value.ToString();
                    GlobalModule.Search2 = Conversions.ToBoolean(ws.Range("A6").Value.ToString());
                    GlobalModule.Flag2 = Conversions.ToBoolean(ws.Range("A7").Value.ToString());
                    GlobalModule.shName2 = ws.Range("A9").Value.ToString();
                    GlobalModule.RangeType2 = ws.Range("B2").value.ToString();
                    // TargetVar = ws.Range("A8").Value.ToString

                    var src_rng = worksheet.get_Range(GlobalModule.GB_CB_Source2);
                    GlobalModule.TType2 = "";

                }
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.Name, "SoftekoSoftekoSofteko", false)))   // For checkbox drop-down list
                {

                    // If worksheet "Neww" is found, store the value of A1 in the variable
                    GlobalModule.GB_CB_Source3 = ws.Range("A2").Value.ToString();

                    GlobalModule.SR3 = ws.Range("A3").Value.ToString();
                    // Horizontal2 = ws.Range("A4").Value.ToString()
                    // Separator2 = ws.Range("A5").Value.ToString()
                    // Search2 = ws.Range("A6").Value.ToString()
                    GlobalModule.Flag3 = Conversions.ToBoolean(ws.Range("A7").Value.ToString());
                    GlobalModule.shName3 = ws.Range("A9").Value.ToString();
                    GlobalModule.RangeType3 = ws.Range("B2").value.ToString();
                    // TargetVar = ws.Range("A8").Value.ToString

                    var src_rng = worksheet.get_Range(GlobalModule.GB_CB_Source3);
                    GlobalModule.TType3 = "";


                }

                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.name, "MySpecialSheet", false)))
                {
                    GlobalModule.Variable1 = ws.Range("A1").Value.ToString();
                    GlobalModule.Variable2 = ws.Range("A2").Value.ToString();
                    GlobalModule.Header = Conversions.ToBoolean(ws.Range("A3").Value.ToString());
                    GlobalModule.Ascending = Conversions.ToBoolean(ws.Range("A4").Value.ToString());
                    GlobalModule.Descending = Conversions.ToBoolean(ws.Range("A5").Value.ToString());
                    GlobalModule.TextConvert = Conversions.ToBoolean(ws.Range("A6").Value.ToString());
                    GlobalModule.OptionType = Conversions.ToBoolean(ws.Range("A7").Value.ToString());
                    GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(ws.Range("A8").Value.ToString());
                    GlobalModule.Flag_CreateDDDL = Conversions.ToBoolean(ws.Range("A9").value.ToString());
                    sheetName3 = ws.Range("A10").value.ToString();
                    sheetName4 = ws.Range("A11").value.ToString();

                }

                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.name, "SoftekoPictureBasedDropDown", false)))
                {
                    GlobalModule.Flag_Picture = Conversions.ToBoolean(ws.Range("A2").value.ToString());
                    GlobalModule.sheetName2 = ws.Range("A3").value.ToString();
                    GlobalModule.Src_Rng_of_PictureDDL = ws.Range("A4").value.ToString();
                    GlobalModule.Des_Rng_of_PictureDDL = ws.Range("A5").value.ToString();

                }
            }
        }

        private void HideNewwwWorksheet()
        {

            // Loop through each worksheet in the active workbook
            foreach (Excel.Worksheet ws in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            {
                if (ws.Name == "Newwwwwwwwww")
                {

                    // If worksheet "Newww" is found, hide it
                    ws.Visible = XlSheetVisibility.xlSheetHidden;
                    // Exit Sub ' No need to check other sheets once "Newww" is found
                }

                if (ws.Name == "SoftekoSofteko")
                {
                    // If worksheet "SoftekoSofteko" is found, hide it

                    ws.Visible = XlSheetVisibility.xlSheetHidden;
                    // Exit Sub ' No need to check other sheets once "Newww" is found
                }

                if (ws.Name == "SoftekoSoftekoSofteko")
                {
                    // If worksheet "SoftekoSofteko" is found, hide it
                    ws.Visible = XlSheetVisibility.xlSheetHidden;
                    // Exit Sub ' No need to check other sheets once "Newww" is found
                }
                if (ws.Name == "SoftekoPictureBasedDropDown")
                {
                    ws.Visible = XlSheetVisibility.xlSheetHidden;
                }
            }
        }

        // For multi drop-down list
        private void sheet_SelectionChange1(Range Target)
        {

            excelApp = Globals.ThisAddIn.Application;
            workBook = excelApp.ActiveWorkbook;
            worksheet = (Excel.Worksheet)workBook.ActiveSheet;

            try
            {


                Range src_rng_concate;


                if (GlobalModule.GB_CB_Source1 is not null)
                {
                    var src_rng = worksheet.get_Range(GlobalModule.GB_CB_Source1);

                    // src_rng = workSheet.Range(GB_CB_Source1)


                    if (GlobalModule.SR1.Contains("Active Workbook"))
                    {
                        src_rng = worksheet.get_Range("A1", worksheet.Cells[excelApp.Rows.Count, excelApp.Columns.Count]);
                    }
                    src_rng = (Range)workBook.ActiveSheet.range(src_rng.get_Address());

                    // Change starts from here
                    if ((GlobalModule.Nam1 ?? "") == (worksheet.Name ?? "") & GlobalModule.TType1 == "Select Range" | (GlobalModule.Nam1 ?? "") == (worksheet.Name ?? "") & GlobalModule.TType1.Contains("Active Sheet") | (GlobalModule.Nam1 ?? "") == (worksheet.Name ?? "") & (GlobalModule.TType1 ?? "") == (worksheet.Name ?? ""))
                    {

                        src_rng_concate = worksheet.get_Range(GlobalModule.GB_CB_Dlt1);


                        if (IsCellInsideRange(Target, src_rng) == true & Target.Cells.Count == 1 & HasDataValidationList(Target) & IsCellInsideRange(Target, src_rng_concate) == true)
                        {
                            // If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                            GlobalModule.TargetVar1 = Target.get_Address();
                            if (Form is not null)
                            {
                                // Form = Nothing

                                Form.Dispose();

                            }
                        }




                        else if (IsCellInsideRange(Target, src_rng) & Target.Cells.Count == 1 & HasDataValidationList(Target))
                        {

                            // If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                            GlobalModule.TargetVar1 = Target.get_Address();
                            if (Form is null || Form.IsDisposed)
                            {
                                Form = new Form36();
                                Form.Show();
                                Form.BringToFront();
                                Form.Refresh();
                            }
                            else
                            {
                                // If form is already open, bring it to the front

                                Form.Dispose();
                                Form = new Form36();
                                Form.Show();
                                Form.BringToFront();

                            }

                            // Dim form As New Form36()
                            // form.Show()
                            // form.Focus()
                            // 'form.TopMost = True
                            // 'form.Activate()
                            // form.BringToFront()
                            // End If
                        }
                    }

                    else if (IsCellInsideRange(Target, src_rng) & Target.Cells.Count == 1 & HasDataValidationList(Target))
                    {
                        // If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                        GlobalModule.TargetVar1 = Target.get_Address();
                        if (Form is null || Form.IsDisposed)
                        {
                            Form = new Form36();
                            Form.Show();
                            Form.BringToFront();
                            Form.Refresh();
                        }
                        else
                        {
                            // If form is already open, bring it to the front

                            Form.Dispose();
                            Form = new Form36();
                            Form.Show();
                            Form.BringToFront();

                        }
                    }
                }
                else
                {

                    // If Form IsNot Nothing Then
                    // 'Form = Nothing

                    // Form.Dispose()

                    // End If


                }
            }

            catch (Exception ex)
            {
                Interaction.MsgBox(ex.Message, MsgBoxStyle.Critical);

            }
        }

        // For CheckBox drop-down List
        private void sheet_SelectionChange2(Range Target)
        {
            excelApp = Globals.ThisAddIn.Application;
            workBook = excelApp.ActiveWorkbook;
            worksheet = (Excel.Worksheet)workBook.ActiveSheet;
            try
            {



                Range src_rng_concate;
                if (GlobalModule.GB_CB_Source2 is not null)
                {
                    var src_rng = worksheet.get_Range(GlobalModule.GB_CB_Source2);

                    if (GlobalModule.SR2.Contains("Active Workbook"))
                    {
                        src_rng = worksheet.get_Range("A1", worksheet.Cells[excelApp.Rows.Count, excelApp.Columns.Count]);
                    }
                    src_rng = (Range)workBook.ActiveSheet.range(src_rng.get_Address());


                    // Change starts from here
                    if ((GlobalModule.Nam2 ?? "") == (worksheet.Name ?? "") & GlobalModule.TType2 == "Select Range" | (GlobalModule.Nam2 ?? "") == (worksheet.Name ?? "") & GlobalModule.TType2.Contains("Active Sheet") | (GlobalModule.Nam2 ?? "") == (worksheet.Name ?? "") & (GlobalModule.TType2 ?? "") == (worksheet.Name ?? ""))
                    {

                        src_rng_concate = worksheet.get_Range(GlobalModule.GB_CB_Dlt2);


                        if (IsCellInsideRange(Target, src_rng) == true & Target.Cells.Count == 1 & HasDataValidationList(Target) & IsCellInsideRange(Target, src_rng_concate) == true)
                        {

                            // If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                            GlobalModule.TargetVar2 = Target.get_Address();
                            if (Form is not null)
                            {
                                // Form = Nothing

                                Form.Dispose();

                            }
                        }



                        else if (IsCellInsideRange(Target, src_rng) & Target.Cells.Count == 1 & HasDataValidationList(Target))
                        {

                            // If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                            GlobalModule.TargetVar2 = Target.get_Address();
                            if (Form2 is null || Form.IsDisposed)
                            {
                                Form2 = new Form38();
                                Form2.Show();
                                Form2.BringToFront();
                                Form2.Refresh();
                            }
                            else
                            {
                                // If form is already open, bring it to the front

                                Form2.Dispose();
                                Form2 = new Form38();
                                Form2.Show();
                                Form2.BringToFront();

                            }

                            // Dim form As New Form36()
                            // form.Show()
                            // form.Focus()
                            // 'form.TopMost = True
                            // 'form.Activate()
                            // form.BringToFront()
                            // End If
                        }
                    }

                    else if (IsCellInsideRange(Target, src_rng) & Target.Cells.Count == 1 & HasDataValidationList(Target))
                    {

                        // If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                        GlobalModule.TargetVar2 = Target.get_Address();
                        if (Form2 is null || Form2.IsDisposed)
                        {
                            Form2 = new Form38();
                            Form2.Show();
                            Form2.BringToFront();
                            Form2.Refresh();
                        }
                        else
                        {
                            // If form is already open, bring it to the front

                            Form2.Dispose();
                            Form2 = new Form38();
                            Form2.Show();
                            Form2.BringToFront();


                        }
                    }
                }
                else
                {

                    // If Form IsNot Nothing Then
                    // 'Form = Nothing

                    // Form.Dispose()

                    // End If


                }
            }

            catch (Exception ex)
            {
                Interaction.MsgBox(ex.Message, MsgBoxStyle.Critical);

            }
        }

        // For search drop-down
        private void sheet_SelectionChange3(Range Target)
        {
            excelApp = Globals.ThisAddIn.Application;
            workBook = excelApp.ActiveWorkbook;
            worksheet = (Excel.Worksheet)workBook.ActiveSheet;

            Range src_rng_concate;


            // src_rng = workSheet.Range(GB_CB_Source1)

            if (GlobalModule.GB_CB_Source3 is not null)
            {
                var src_rng = worksheet.get_Range(GlobalModule.GB_CB_Source3);

                // src_rng = workSheet.Range(GB_CB_Source1)


                if (GlobalModule.SR3.Contains("Active Workbook"))
                {
                    src_rng = worksheet.get_Range("A1", worksheet.Cells[excelApp.Rows.Count, excelApp.Columns.Count]);
                }
                src_rng = (Range)workBook.ActiveSheet.range(src_rng.get_Address());

                // Change starts from here
                if ((GlobalModule.Nam3 ?? "") == (worksheet.Name ?? "") & GlobalModule.TType3 == "Select Range" | (GlobalModule.Nam3 ?? "") == (worksheet.Name ?? "") & GlobalModule.TType3.Contains("Active Sheet") | (GlobalModule.Nam3 ?? "") == (worksheet.Name ?? "") & (GlobalModule.TType3 ?? "") == (worksheet.Name ?? ""))
                {

                    src_rng_concate = worksheet.get_Range(GlobalModule.GB_CB_Dlt3);

                    if (IsCellInsideRange(Target, src_rng) == true & Target.Cells.Count == 1 & HasDataValidationList(Target) & IsCellInsideRange(Target, src_rng_concate) == true)
                    {

                        // If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                        GlobalModule.TargetVar3 = Target.get_Address();
                        if (Form3 is not null)
                        {
                            // Form = Nothing

                            Form3.Dispose();

                        }
                    }




                    else if (IsCellInsideRange(Target, src_rng) & Target.Cells.Count == 1 & HasDataValidationList(Target))
                    {

                        // If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                        GlobalModule.TargetVar3 = Target.get_Address();
                        if (Form3 is null || Form.IsDisposed)
                        {
                            Form3 = new Form40();
                            Form3.Show();
                            Form3.BringToFront();
                            Form3.Refresh();
                        }
                        else
                        {
                            // If form is already open, bring it to the front

                            Form3.Dispose();
                            Form3 = new Form40();
                            Form3.Show();
                            Form3.BringToFront();

                        }

                        // Dim form As New Form36()
                        // form.Show()
                        // form.Focus()
                        // 'form.TopMost = True
                        // 'form.Activate()
                        // form.BringToFront()
                        // End If
                    }
                }

                else if (IsCellInsideRange(Target, src_rng) & Target.Cells.Count == 1 & HasDataValidationList(Target))
                {

                    // If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                    GlobalModule.TargetVar3 = Target.get_Address();
                    if (Form3 is null || Form.IsDisposed)
                    {
                        Form3 = new Form40();
                        Form3.Show();
                        Form3.BringToFront();
                        Form3.Refresh();
                    }
                    else
                    {
                        // If form is already open, bring it to the front

                        Form3.Dispose();
                        Form3 = new Form40();
                        Form3.Show();
                        Form3.BringToFront();


                    }
                }
            }
            else
            {

                // If Form IsNot Nothing Then
                // 'Form = Nothing

                // Form.Dispose()

                // End If


            }
        }
        private bool IsCellInsideRange(Range cell, Range targetRange)
        {

            try
            {
                var intersectRange = Globals.ThisAddIn.Application.Intersect(cell, targetRange);

                return intersectRange is not null;
            }
            catch (Exception ex)
            {

                return false;
            }
        }

        private bool HasDataValidationList(Range cell)
        {
            bool hasValidation = false;

            try
            {
                if (cell.Validation is not null && cell.Validation.Type == (int)XlDVType.xlValidateList)
                {
                    hasValidation = true;
                }
            }
            catch (Exception ex)
            {
                // Exception will be thrown if cell doesn't have validation. No action needed.
            }

            return hasValidation;
        }


        // Create Dynamic Drop-down list
        public void worksheet5_1_Change(Range Target)
        {
            excelApp = Globals.ThisAddIn.Application;
            workBook = excelApp.ActiveWorkbook;
            worksheet = (Excel.Worksheet)workBook.ActiveSheet;

            // src_rng = excelApp.Range(Variable1)
            // des_rng = excelApp.Range(Variable2)
            // Dim src_sheet As Excel.Worksheet = CType(workBook.Worksheets(sheetName3), Excel.Worksheet)
            // Dim des_sheet As Excel.Worksheet = CType(workBook.Worksheets(sheetName4), Excel.Worksheet)

            // src_rng = src_sheet.Range(src_rng.Address)

            // des_rng = des_sheet.Range(des_rng.Address)

            Range rng;
            // des_rng.ClearContents()

            if (GlobalModule.OptionType == true)
            {
                if (GlobalModule.Header == true)
                {
                    // Dim adjustRange As Excel.Range
                    rng = src_rng.get_Offset(1, 0).get_Resize(src_rng.Rows.Count - 1, src_rng.Columns.Count);
                }

                else
                {

                    rng = src_rng;
                } // Assuming you have a range from A1 to A100

                int col_dif;
                col_dif = Target.Column - worksheet.get_Range(des_rng.get_Address()).Column + 1;


                // For k = 1 To des_rng.Rows.Count
                var matchedValues = new List<string>();
                var sec_matchedValues = new List<string>();
                var thrd_matchedValues = new List<string>();
                var four_matchedValues = new List<string>();
                int k = Target.Row - worksheet.get_Range(des_rng.get_Address()).Row + 1;

                if (col_dif == 1)
                {

                    if (des_rng[k, 1].Value is not null)
                    {
                        for (int i = 1, loopTo = rng.Rows.Count; i <= loopTo; i++)
                        {
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(rng[i, 1].Value, des_rng[k, 1].Value, false)))
                            {
                                if (!matchedValues.Contains(Conversions.ToString(rng[i, 2].Value)))
                                {
                                    matchedValues.Add(Conversions.ToString(rng[i, 2].Value));
                                }
                                // matchedValues.Add(rng(i, 2).Value)
                            }
                        }


                        if (GlobalModule.Ascending == true)
                        {
                            // Sort the list in ascending order
                            matchedValues.Sort();
                        }
                        else if (GlobalModule.Descending == true)
                        {
                            // Sort the list in ascending order
                            matchedValues.Sort();
                            matchedValues.Reverse();
                        }

                        // Dim dropDownRange As Excel.Range = des_rng(k, 2)
                        Range dropDownRange = (Range)Target[1, 2];

                        var Validation = dropDownRange.Validation;
                        Validation.Delete(); // Remove any existing validation
                        Validation.Add(XlDVType.xlValidateList, Formula1: string.Join(",", matchedValues));
                        matchedValues.Clear();

                    }
                }

                // Dim sec_matchedValues As New List(Of String)
                else if (col_dif == 2)
                {
                    if (des_rng[k, 2].Value is not null)
                    {
                        for (int i = 1, loopTo1 = rng.Rows.Count; i <= loopTo1; i++)
                        {
                            if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(rng[i, 1].Value, des_rng[k, 1].Value, false), Operators.ConditionalCompareObjectEqual(rng[i, 2].Value, des_rng[k, 2].Value, false))))
                            {
                                if (!sec_matchedValues.Contains(Conversions.ToString(rng[i, 3].Value)))
                                {
                                    sec_matchedValues.Add(Conversions.ToString(rng[i, 3].Value));
                                }

                            }
                        }


                        if (GlobalModule.Ascending == true)
                        {
                            // Sort the list in ascending order
                            sec_matchedValues.Sort();
                        }
                        else if (GlobalModule.Descending == true)
                        {
                            // Sort the list in ascending order
                            sec_matchedValues.Sort();
                            sec_matchedValues.Reverse();
                        }


                        // Dim dropDownRange As Excel.Range = des_rng(k, 3)
                        Range dropDownRange = default;
                        var Validation = dropDownRange.Validation;
                        Validation.Delete(); // Remove any existing validation
                        Validation.Add(XlDVType.xlValidateList, Formula1: string.Join(",", sec_matchedValues));
                        sec_matchedValues.Clear();
                    }
                }
                else if (col_dif == 3)
                {
                    // Dim thrd_matchedValues As New List(Of String)

                    if (des_rng[k, 3].Value is not null)
                    {
                        for (int i = 1, loopTo2 = rng.Rows.Count; i <= loopTo2; i++)
                        {
                            if (Conversions.ToBoolean(Operators.AndObject(Operators.AndObject(Operators.ConditionalCompareObjectEqual(rng[i, 1].Value, des_rng[k, 1].Value, false), Operators.ConditionalCompareObjectEqual(rng[i, 2].Value, des_rng[k, 2].Value, false)), Operators.ConditionalCompareObjectEqual(rng[i, 3].Value, des_rng[k, 3].Value, false))))
                            {
                                if (!thrd_matchedValues.Contains(Conversions.ToString(rng[i, 4].Value)))
                                {
                                    thrd_matchedValues.Add(Conversions.ToString(rng[i, 4].Value));
                                }

                            }
                        }


                        if (GlobalModule.Ascending == true)
                        {
                            // Sort the list in ascending order
                            thrd_matchedValues.Sort();
                        }
                        else if (GlobalModule.Descending == true)
                        {
                            // Sort the list in ascending order
                            thrd_matchedValues.Sort();
                            thrd_matchedValues.Reverse();
                        }


                        // Dim dropDownRange As Excel.Range = des_rng(k, 4)
                        Range dropDownRange = default;
                        var Validation = dropDownRange.Validation;
                        Validation.Delete(); // Remove any existing validation
                        Validation.Add(XlDVType.xlValidateList, Formula1: string.Join(",", thrd_matchedValues));
                        thrd_matchedValues.Clear();
                    }
                }


                // Dim four_matchedValues As New List(Of String)
                else if (col_dif == 4)
                {
                    if (des_rng[k, 4].Value is not null)
                    {
                        for (int i = 1, loopTo3 = rng.Rows.Count; i <= loopTo3; i++)
                        {
                            if (Conversions.ToBoolean(Operators.AndObject(Operators.AndObject(Operators.AndObject(Operators.ConditionalCompareObjectEqual(rng[i, 1].Value, des_rng[k, 1].Value, false), Operators.ConditionalCompareObjectEqual(rng[i, 2].Value, des_rng[k, 2].Value, false)), Operators.ConditionalCompareObjectEqual(rng[i, 3].Value, des_rng[k, 3].Value, false)), Operators.ConditionalCompareObjectEqual(rng[i, 4].Value, des_rng[k, 4].Value, false))))
                            {

                                if (!four_matchedValues.Contains(Conversions.ToString(rng[i, 5].Value)))
                                {
                                    four_matchedValues.Add(Conversions.ToString(rng[i, 5].Value));
                                }


                            }
                        }


                        if (GlobalModule.Ascending == true)
                        {
                            // Sort the list in ascending order
                            four_matchedValues.Sort();
                        }
                        else if (GlobalModule.Descending == true)
                        {
                            // Sort the list in ascending order
                            four_matchedValues.Sort();
                            four_matchedValues.Reverse();
                        }


                        Range dropDownRange = (Range)des_rng[k, 5];
                        var Validation = dropDownRange.Validation;
                        Validation.Delete(); // Remove any existing validation
                        Validation.Add(XlDVType.xlValidateList, Formula1: string.Join(",", four_matchedValues));
                        four_matchedValues.Clear();
                    }
                }
            }

            // Next

            else if (GlobalModule.OptionType == false)
            {
                if (GlobalModule.Horizontal_CreateDP == true)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Target.get_Address(), des_rng[1, 1].Address, false)))
                    {

                        var worksheet = Target.Worksheet;
                        int col = src_rng.Rows.Find(Target.get_Value()).Column - src_rng.Column + 1;

                        // Dim ab As Integer = col - src_rng.Column
                        Range sourceRng = (Range)src_rng.Cells[2, col].Resize(Operators.SubtractObject(src_rng[src_rng.Rows.Count, col].row, 2), (object)1);

                        // Dim sourceRng As Excel.Range = src_rng.Cells(2, col).Resize(worksheet.Cells(worksheet.Rows.Count, col), 1)
                        Range dropDownRange = (Range)des_rng[1, 2];
                        var Validation = dropDownRange.Validation;
                        Validation.Delete(); // Remove any existing validation
                        string formula = "='" + sourceRng.Worksheet.Name + "'!" + sourceRng.get_Address(External: false);
                        Validation.Add(XlDVType.xlValidateList, Formula1: formula);
                        // CreateValidationList(worksheet.Cells(2, 5), "=" & sourceRng.Address)
                    }
                }

                else if (GlobalModule.Horizontal_CreateDP == false)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Target.get_Address(), des_rng[1, 1].Address, false)))
                    {
                        var worksheet = Target.Worksheet;
                        int col = src_rng.Rows.Find(Target.get_Value()).Column - src_rng.Column + 1;

                        // Dim ab As Integer = col - src_rng.Column
                        Range sourceRng = (Range)src_rng.Cells[2, col].Resize(Operators.SubtractObject(src_rng[src_rng.Rows.Count, col].row, 2), (object)1);

                        // Dim sourceRng As Excel.Range = src_rng.Cells(2, col).Resize(worksheet.Cells(worksheet.Rows.Count, col), 1)
                        Range dropDownRange = (Range)des_rng[2, 1];
                        var Validation = dropDownRange.Validation;
                        Validation.Delete(); // Remove any existing validation
                        string formula = "='" + sourceRng.Worksheet.Name + "'!" + sourceRng.get_Address(External: false);
                        Validation.Add(XlDVType.xlValidateList, Formula1: formula);
                    }
                }

            }
            // Catch ex As Exception

            // End Try


        }


        public void worksheet5_2_Change(Range Target)
        {
            excelApp = Globals.ThisAddIn.Application;
            workBook = excelApp.ActiveWorkbook;
            worksheet = (Excel.Worksheet)workBook.ActiveSheet;

            var targetWorksheet = default(Excel.Worksheet);
            int i = 1;
            bool j = false;
            foreach (var ws in excelApp.ActiveWorkbook.Worksheets)
            {
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.name, "MySpecialSheet", false)))
                {
                    targetWorksheet = (Excel.Worksheet)ws;
                    j = true;
                    break;
                }
            }
            if (j == true)
            {
                Range r11 = null;
                Range r12 = null;
                Range r13 = null;
                Range r14 = null;
                Range r15 = null;

                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(targetWorksheet.get_Range("A1").get_Value(), "", false)))
                {
                    r11 = excelApp.get_Range(targetWorksheet.get_Range("A2").get_Value());
                    r11 = worksheet.get_Range(r11.get_Address());
                    // MsgBox(r11.Worksheet.Name)
                }

                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(targetWorksheet.get_Range("B1").get_Value(), "", false)))
                {
                    r12 = excelApp.get_Range(targetWorksheet.get_Range("B2").get_Value());
                    r12 = worksheet.get_Range(r12.get_Address());
                }

                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(targetWorksheet.get_Range("C1").get_Value(), "", false)))
                {
                    r13 = excelApp.get_Range(targetWorksheet.get_Range("C2").get_Value());
                    r13 = worksheet.get_Range(r13.get_Address());
                }

                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(targetWorksheet.get_Range("D1").get_Value(), "", false)))
                {
                    r14 = excelApp.get_Range(targetWorksheet.get_Range("D2").get_Value());
                    r14 = worksheet.get_Range(r14.get_Address());
                }

                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(targetWorksheet.get_Range("E1").get_Value(), "", false)))
                {
                    r15 = excelApp.get_Range(targetWorksheet.get_Range("E2").get_Value());
                    r15 = worksheet.get_Range(r15.get_Address());
                    // MsgBox(r15.Address)
                    // MsgBox(excelApp.Intersect(Target, r15).Address)
                    // MsgBox(Target.Worksheet.Name)
                    // MsgBox(targetWorksheet.Range("E11").Value)
                }

                // For i = 1 To targetWorksheet.Columns.Count
                if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(Target.Worksheet.Name, targetWorksheet.get_Range("A11").get_Value(), false), excelApp.Intersect(Target, r11) is not null)))
                {
                    // If excelApp.Intersect(Target, r11) IsNot Nothing Then
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
                }
                // End If
                else if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(Target.Worksheet.Name, targetWorksheet.get_Range("B11").get_Value(), false), excelApp.Intersect(Target, r12) is not null)))
                {
                    // If excelApp.Intersect(Target, r12) IsNot Nothing Then
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
                }
                // End If

                else if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(Target.Worksheet.Name, targetWorksheet.get_Range("C11").get_Value(), false), excelApp.Intersect(Target, r13) is not null)))
                {
                    // If excelApp.Intersect(Target, r13) IsNot Nothing Then
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
                }
                // End If

                else if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(Target.Worksheet.Name, targetWorksheet.get_Range("D11").get_Value(), false), excelApp.Intersect(Target, r14) is not null)))
                {
                    // If excelApp.Intersect(Target, r14) IsNot Nothing Then
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
                }
                // End If

                else if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(Target.Worksheet.Name, targetWorksheet.get_Range("E11").get_Value(), false), excelApp.Intersect(Target, r15) is not null)))
                {
                    // If excelApp.Intersect(Target, r15) IsNot Nothing Then
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
                    // End If
                }

                src_rng = excelApp.get_Range(GlobalModule.Variable1);
                Excel.Worksheet src_ws = (Excel.Worksheet)workBook.Worksheets[GlobalModule.sheetName10];
                Excel.Worksheet des_ws = (Excel.Worksheet)workBook.Worksheets[GlobalModule.sheetName11];
                src_rng = src_ws.get_Range(GlobalModule.Variable1);


                // des_rng = des_ws.Range(des_rng.Address)
                des_rng = des_ws.get_Range(GlobalModule.Variable2);


                if (excelApp.Intersect(Target, des_rng) is not null)
                {


                    Range rng;

                    // Dim rng As Excel.Range
                    // des_rng.ClearContents()

                    if (GlobalModule.OptionType == true)
                    {
                        if (GlobalModule.Header == true)
                        {
                            // Dim adjustRange As Excel.Range
                            rng = src_rng.get_Offset(1, 0).get_Resize(src_rng.Rows.Count - 1, src_rng.Columns.Count);
                        }

                        else
                        {

                            rng = src_rng;
                        } // Assuming you have a range from A1 to A100

                        int col_dif;
                        col_dif = Target.Column - worksheet.get_Range(des_rng.get_Address()).Column + 1;


                        // For k = 1 To des_rng.Rows.Count
                        var matchedValues = new List<string>();
                        var sec_matchedValues = new List<string>();
                        var thrd_matchedValues = new List<string>();
                        var four_matchedValues = new List<string>();
                        int k = Target.Row - worksheet.get_Range(des_rng.get_Address()).Row + 1;

                        if (col_dif == 1)
                        {

                            if (des_rng[k, 1].Value is not null)
                            {
                                var loopTo = rng.Rows.Count;
                                for (i = 1; i <= loopTo; i++)
                                {
                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(rng[i, 1].Value, des_rng[k, 1].Value, false)))
                                    {
                                        if (!matchedValues.Contains(Conversions.ToString(rng[i, 2].Value)))
                                        {
                                            matchedValues.Add(Conversions.ToString(rng[i, 2].Value));
                                        }
                                        // matchedValues.Add(rng(i, 2).Value)
                                    }
                                }


                                if (GlobalModule.Ascending == true)
                                {
                                    // Sort the list in ascending order
                                    matchedValues.Sort();
                                }
                                else if (GlobalModule.Descending == true)
                                {
                                    // Sort the list in ascending order
                                    matchedValues.Sort();
                                    matchedValues.Reverse();
                                }

                                // Dim dropDownRange As Excel.Range = des_rng(k, 2)
                                Range dropDownRange = (Range)Target[1, 2];

                                var Validation = dropDownRange.Validation;
                                Validation.Delete(); // Remove any existing validation
                                Validation.Add(XlDVType.xlValidateList, Formula1: string.Join(",", matchedValues));
                                matchedValues.Clear();

                            }
                        }

                        // Dim sec_matchedValues As New List(Of String)
                        else if (col_dif == 2)
                        {
                            if (des_rng[k, 2].Value is not null)
                            {
                                var loopTo1 = rng.Rows.Count;
                                for (i = 1; i <= loopTo1; i++)
                                {
                                    if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(rng[i, 1].Value, des_rng[k, 1].Value, false), Operators.ConditionalCompareObjectEqual(rng[i, 2].Value, des_rng[k, 2].Value, false))))
                                    {
                                        if (!sec_matchedValues.Contains(Conversions.ToString(rng[i, 3].Value)))
                                        {
                                            sec_matchedValues.Add(Conversions.ToString(rng[i, 3].Value));
                                        }

                                    }
                                }


                                if (GlobalModule.Ascending == true)
                                {
                                    // Sort the list in ascending order
                                    sec_matchedValues.Sort();
                                }
                                else if (GlobalModule.Descending == true)
                                {
                                    // Sort the list in ascending order
                                    sec_matchedValues.Sort();
                                    sec_matchedValues.Reverse();
                                }


                                // Dim dropDownRange As Excel.Range = des_rng(k, 3)
                                Range dropDownRange = default;
                                var Validation = dropDownRange.Validation;
                                Validation.Delete(); // Remove any existing validation
                                Validation.Add(XlDVType.xlValidateList, Formula1: string.Join(",", sec_matchedValues));
                                sec_matchedValues.Clear();
                            }
                        }
                        else if (col_dif == 3)
                        {
                            // Dim thrd_matchedValues As New List(Of String)

                            if (des_rng[k, 3].Value is not null)
                            {
                                var loopTo2 = rng.Rows.Count;
                                for (i = 1; i <= loopTo2; i++)
                                {
                                    if (Conversions.ToBoolean(Operators.AndObject(Operators.AndObject(Operators.ConditionalCompareObjectEqual(rng[i, 1].Value, des_rng[k, 1].Value, false), Operators.ConditionalCompareObjectEqual(rng[i, 2].Value, des_rng[k, 2].Value, false)), Operators.ConditionalCompareObjectEqual(rng[i, 3].Value, des_rng[k, 3].Value, false))))
                                    {
                                        if (!thrd_matchedValues.Contains(Conversions.ToString(rng[i, 4].Value)))
                                        {
                                            thrd_matchedValues.Add(Conversions.ToString(rng[i, 4].Value));
                                        }

                                    }
                                }


                                if (GlobalModule.Ascending == true)
                                {
                                    // Sort the list in ascending order
                                    thrd_matchedValues.Sort();
                                }
                                else if (GlobalModule.Descending == true)
                                {
                                    // Sort the list in ascending order
                                    thrd_matchedValues.Sort();
                                    thrd_matchedValues.Reverse();
                                }


                                // Dim dropDownRange As Excel.Range = des_rng(k, 4)
                                Range dropDownRange = default;
                                var Validation = dropDownRange.Validation;
                                Validation.Delete(); // Remove any existing validation
                                Validation.Add(XlDVType.xlValidateList, Formula1: string.Join(",", thrd_matchedValues));
                                thrd_matchedValues.Clear();
                            }
                        }


                        // Dim four_matchedValues As New List(Of String)
                        else if (col_dif == 4)
                        {
                            if (des_rng[k, 4].Value is not null)
                            {
                                var loopTo3 = rng.Rows.Count;
                                for (i = 1; i <= loopTo3; i++)
                                {
                                    if (Conversions.ToBoolean(Operators.AndObject(Operators.AndObject(Operators.AndObject(Operators.ConditionalCompareObjectEqual(rng[i, 1].Value, des_rng[k, 1].Value, false), Operators.ConditionalCompareObjectEqual(rng[i, 2].Value, des_rng[k, 2].Value, false)), Operators.ConditionalCompareObjectEqual(rng[i, 3].Value, des_rng[k, 3].Value, false)), Operators.ConditionalCompareObjectEqual(rng[i, 4].Value, des_rng[k, 4].Value, false))))
                                    {

                                        if (!four_matchedValues.Contains(Conversions.ToString(rng[i, 5].Value)))
                                        {
                                            four_matchedValues.Add(Conversions.ToString(rng[i, 5].Value));
                                        }


                                    }
                                }


                                if (GlobalModule.Ascending == true)
                                {
                                    // Sort the list in ascending order
                                    four_matchedValues.Sort();
                                }
                                else if (GlobalModule.Descending == true)
                                {
                                    // Sort the list in ascending order
                                    four_matchedValues.Sort();
                                    four_matchedValues.Reverse();
                                }


                                Range dropDownRange = (Range)des_rng[k, 5];
                                var Validation = dropDownRange.Validation;
                                Validation.Delete(); // Remove any existing validation
                                Validation.Add(XlDVType.xlValidateList, Formula1: string.Join(",", four_matchedValues));
                                four_matchedValues.Clear();
                            }
                        }
                    }

                    // Next

                    else if (GlobalModule.OptionType == false)
                    {
                        if (GlobalModule.Horizontal_CreateDP == true)
                        {
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Target.get_Address(), des_rng[1, 1].Address, false)))
                            {

                                var worksheet = Target.Worksheet;
                                int col = src_rng.Rows.Find(Target.get_Value()).Column - src_rng.Column + 1;

                                // Dim ab As Integer = col - src_rng.Column
                                Range sourceRng = (Range)src_rng.Cells[2, col].Resize(Operators.SubtractObject(src_rng[src_rng.Rows.Count, col].row, 2), (object)1);

                                // Dim sourceRng As Excel.Range = src_rng.Cells(2, col).Resize(worksheet.Cells(worksheet.Rows.Count, col), 1)
                                Range dropDownRange = (Range)des_rng[1, 2];
                                var Validation = dropDownRange.Validation;
                                Validation.Delete(); // Remove any existing validation
                                string formula = "='" + sourceRng.Worksheet.Name + "'!" + sourceRng.get_Address(External: false);
                                Validation.Add(XlDVType.xlValidateList, Formula1: formula);
                                // CreateValidationList(worksheet.Cells(2, 5), "=" & sourceRng.Address)
                            }
                        }

                        else if (GlobalModule.Horizontal_CreateDP == false)
                        {
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Target.get_Address(), des_rng[1, 1].Address, false)))
                            {
                                var worksheet = Target.Worksheet;
                                int col = src_rng.Rows.Find(Target.get_Value()).Column - src_rng.Column + 1;

                                // Dim ab As Integer = col - src_rng.Column
                                Range sourceRng = (Range)src_rng.Cells[2, col].Resize(Operators.SubtractObject(src_rng[src_rng.Rows.Count, col].row, 2), (object)1);

                                // Dim sourceRng As Excel.Range = src_rng.Cells(2, col).Resize(worksheet.Cells(worksheet.Rows.Count, col), 1)
                                Range dropDownRange = (Range)des_rng[2, 1];
                                var Validation = dropDownRange.Validation;
                                Validation.Delete(); // Remove any existing validation
                                string formula = "='" + sourceRng.Worksheet.Name + "'!" + sourceRng.get_Address(External: false);
                                Validation.Add(XlDVType.xlValidateList, Formula1: formula);
                            }
                        }

                    }


                }
            }
            // Catch ex As Exception
            // MsgBox("error")
            // End Try
            // MsgBox(3)

        }

        // For picturebox Drop-down
        private void worksheet6_Change(Range Target)
        {

            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            var src_rng = excelApp.get_Range(GlobalModule.Src_Rng_of_PictureDDL);

            // Target = worksheet.Range(Target.Address)

            for (int i = 1, loopTo = src_rng.Rows.Count; i <= loopTo; i++)
            {
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(src_rng[i, 1].Value, Target.get_Value(), false)))
                {

                    try
                    {
                        worksheet7_Change(Target);
                    }

                    catch (Exception ex)
                    {

                    }



                    // Dim imageCell As Excel.Range = worksheet.Range(src_rng(i, 2).address)
                    // imageCell.CopyPicture(
                    // Appearance:=Excel.XlPictureAppearance.xlScreen,
                    // Format:=Excel.XlCopyPictureFormat.xlPicture)
                    // worksheet.Paste(Target.Offset(0, 1))
                    // Me.Refresh()


                    bool x = false;

                    foreach (Shape pic in worksheet.Shapes)
                    {

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(pic.TopLeftCell.get_Address(), src_rng[i, 2].Address, false)))
                        {

                            pic.CopyPicture();
                            worksheet.Paste(Target.get_Offset(0, 1));
                            Target.get_Offset(0, 1).RowHeight = src_rng[i, 2].RowHeight;
                            // Target.Offset(0, 1).RowHeight = src_rng(i, 2).C
                            x = true;
                            break;
                        }
                        // x = x + 1
                    }

                    excelApp.CutCopyMode = (XlCutCopyMode)Conversions.ToInteger(false);
                    // Exit Sub

                }
            }


        }

        private void worksheet7_Change(Range Target)
        {

            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;


            foreach (Shape pic in worksheet.Shapes)
            {

                if ((pic.TopLeftCell.get_Address() ?? "") == (Target.get_Offset(0, 1).get_Address() ?? ""))
                {

                    pic.Delete();
                    // Exit For
                }
            }

            // End Sub
        }


    }
}