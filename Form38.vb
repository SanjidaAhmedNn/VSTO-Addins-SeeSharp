Imports System.Data
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel
Imports DataTable = System.Data.DataTable
Imports Point = System.Drawing.Point

Public Class Form38
    Private dt As DataTable
    'Public dv As DataView
    'dim dv As New DataView(dt)

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Public Shared workSheet As Excel.Worksheet

    Dim src_rng As Excel.Range
    Public des_rng As Excel.Range
    Dim selectedRange As Excel.Range

    Public validationRange As Excel.Range
    Public form As Form37_MSDropDownCheckBox = Nothing

    Private processingEvent As Boolean = False
    'Public focuschange As Boolean

    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1

    'Public Target As Excel.Range


    Private Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowRect(ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function


    Private Sub Form38_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Separator = ","
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet


        ' Add Increment Button
        Dim incrementColumn As New DataGridViewCheckBoxColumn


        incrementColumn.Name = "Increment"
        incrementColumn.HeaderText = "Check"

        incrementColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        incrementColumn.Width = 45 ' Set the width you want here
        'incrementColumn.DefaultCellStyle.BackColor = Color.White
        incrementColumn.FlatStyle = FlatStyle.Popup

        incrementColumn.DefaultCellStyle.Font = New System.Drawing.Font("Segoe UI", 12)
        incrementColumn.ReadOnly = False

        '.Columns(1).ReadOnly = True
        DataGridView1.Columns.Add(incrementColumn)


        'Dim formula1 As String
        'Dim dropdownItems() As String
        Dim sourceRange As Excel.Range

        ' Get the cell with the drop-down list
        Dim cell As Excel.Range = worksheet.Range(TargetVar2)


        ' Populate DataGridView
        dt = New DataTable()
        'dt.Columns.Add("Value", GetType(String))
        Dim connectedstring As New List(Of String)()
        Dim connectedstringst As String = ""



        Dim validationList As String = ""
        Dim formula As String = cell.Validation.Formula1
        If formula.Contains(",") Then
            ' Data validation type: Excel Range
            validationList = formula

            dt.Columns.Add("Value", GetType(String))
            dt.Rows.Add("Select all")
            Dim items As String() = validationList.Split(","c)
            For Each item As String In items
                dt.Rows.Add(item)
            Next
        Else
            ' Data validation type: Excel Range

            sourceRange = worksheet.Range(formula)

            dt.Columns.Add("Value", GetType(String))

            dt.Rows.Add("Select all")
            'Sample Data
            For Each itemCell As Excel.Range In sourceRange
                dt.Rows.Add(itemCell.Value)
                'dt.Rows.Add(20)
                'dt.Rows.Add(30)
            Next

        End If


        DataGridView1.DataSource = dt


        If cell.Value IsNot Nothing Then
            connectedstringst = cell.Value.ToString
        End If


        'MsgBox(connectedstringst)

        ' Parse the values
        ' Dim values As String() = worksheet.Range(TargetVar).Split(","c)

        ' DataGridView1.DataSource = dt
        Dim i As Integer = 0

        If connectedstringst IsNot Nothing Or connectedstringst <> "" Then

            For Each r As DataGridViewRow In DataGridView1.Rows

                Dim cellValue As String = r.Cells(1).Value.ToString()

                If connectedstringst.Contains(cellValue.ToString) Then
                    DataGridView1.Rows(i).Cells("Increment").Value = True
                Else
                    DataGridView1.Rows(i).Cells("Increment").Value = False
                End If
                i = i + 1
            Next
        End If

        'MsgBox(connectedstringst)

        'For Each r As DataGridViewRow In DataGridView1.Rows
        '    If Not r.IsNewRow Then ' Avoid the last empty row
        '        'MsgBox(1)

        '        Dim cellValue As Object = r.Cells(1).Value
        '        If cellValue IsNot Nothing AndAlso connectedstringst.Contains(cellValue.ToString()) Then
        '            r.Cells("Increment").Value = True
        '        Else
        '            r.Cells("Increment").Value = False
        '        End If
        '    End If
        'Next



        DataGridView1.Columns(1).Width = 190
        Dim targetCellValue As String = worksheet.Range(TargetVar2).Value 'Assuming B1 cell contains the values
        'Me.BringToFront()
        Me.TopMost = True
        'Me.Show()
        Me.TopMost = False

    End Sub


    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
        ' Ensure it's not the header row
        If e.RowIndex < 0 Then Return

        Dim cell As DataGridViewCell = DataGridView1.Rows(e.RowIndex).Cells("Value")


        '' Check for Increment button click
        'If e.ColumnIndex = DataGridView1.Columns("Increment").Index Then
        '    cell.Value = Convert.ToInt32(cell.Value) + 1
        'End If

        '' Check for Decrement button click
        'If e.ColumnIndex = DataGridView1.Columns("Decrement").Index Then
        '    cell.Value = Convert.ToInt32(cell.Value) - 1
        'End If
        If e.ColumnIndex = DataGridView1.Columns("Increment").Index And DataGridView1.Rows(e.RowIndex).Cells("Increment").Value = False And DataGridView1.Rows(e.RowIndex).Cells(1).Value = "Select all" Then
            'DataGridView1.Rows(e.RowIndex).Cells("Increment").Value = True



            For Each r As DataGridViewRow In DataGridView1.Rows



                'DataGridView1.Rows(i).Cells("Increment").Value = True
                r.Cells(0).Value = True
                'DataGridView1.Rows(j).Cells("Increment") = True
                'MsgBox(1)


            Next




        ElseIf e.ColumnIndex = DataGridView1.Columns("Increment").Index And DataGridView1.Rows(e.RowIndex).Cells("Increment").Value = True And DataGridView1.Rows(e.RowIndex).Cells(1).Value = "Select all" Then
            'DataGridView1.Rows(e.RowIndex).Cells("Increment").Value = False



            For Each r As DataGridViewRow In DataGridView1.Rows

                ' Dim cellValue As String = r.Cells(1).Value.ToString()

                'DataGridView1.Rows(i).Cells("Increment").Value = True
                r.Cells(0).Value = False
                'DataGridView1.Rows(j).Cells("Increment") = True
                'MsgBox(1)


            Next
            'End If

        ElseIf e.ColumnIndex = DataGridView1.Columns("Increment").Index And DataGridView1.Rows(e.RowIndex).Cells("Increment").Value = False And DataGridView1.Rows(e.RowIndex).Cells(1).Value <> "Select all" Then


        DataGridView1.Rows(e.RowIndex).Cells("Increment").Value = True 'or False
            If worksheet.Range(TargetVar2).Value Is Nothing Then
                worksheet.Range(TargetVar2).Value = cell.Value
            Else
                Dim itemToRemove As String = cell.Value


            End If

        ElseIf e.ColumnIndex = DataGridView1.Columns("Increment").Index And DataGridView1.Rows(e.RowIndex).Cells("Increment").Value = True And DataGridView1.Rows(e.RowIndex).Cells(1).Value <> "Select all" Then
            DataGridView1.Rows(e.RowIndex).Cells("Increment").Value = False
            'Me.Refresh()
            Dim itemToRemove As String = cell.Value

            If worksheet.Range(TargetVar2).Value IsNot Nothing AndAlso worksheet.Range(TargetVar2).Value.ToString().Contains(itemToRemove) Then



                Dim items As List(Of String) = worksheet.Range(TargetVar2).Value.ToString().Split(New String() {Separator2}, StringSplitOptions.None).ToList()

                ' Find the index of the first occurrence of the item to remove
                Dim indexToRemove As Integer = items.FindIndex(Function(x) x.Trim() = itemToRemove)

                If indexToRemove >= 0 Then ' If found
                    items.RemoveAt(indexToRemove) ' Remove only the first occurrence
                    worksheet.Range(TargetVar2).Value = String.Join(Separator2, items)
                End If
            End If

        End If




    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs)
        'Dim searchTerm As String = txtSearch.Text.Trim()
        'If String.IsNullOrEmpty(searchTerm) Then
        '    DataGridView1.DataSource = dt
        'Else
        '    Dim dv As New DataView(dt)

        '    If IsNumeric(searchTerm) Then
        '        dv.RowFilter = String.Format("Value = {0}", Convert.ToInt32(searchTerm))
        '        DataGridView1.DataSource = dv
        '    Else
        '        'MessageBox.Show("Please enter a valid number.")
        '    End If
        'End If
    End Sub

    'dv.RowFilter = String.Format("Convert(Value, 'System.String') LIKE '{0}%'", searchTerm)

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged

        Dim searchTerm As String = txtSearch.Text.Trim()

        If String.IsNullOrEmpty(searchTerm) Then
            DataGridView1.DataSource = dt
        Else
            Dim dv As New DataView(dt)

            If IsNumeric(searchTerm) Then
                dv.RowFilter = String.Format("Convert(Value, 'System.String') LIKE '{0}%'", searchTerm)
                DataGridView1.DataSource = dv
            Else
                DataGridView1.DataSource = dt
            End If
        End If


        'Dim searchTerm As String = txtSearch.Text.Trim()
        ''Dim dv As New DataView(dt)

        'If String.IsNullOrEmpty(searchTerm) Then

        '    dv.RowFilter = "" ' Clear the filter
        'Else
        '    If IsNumeric(searchTerm) Then
        '        dv.RowFilter = String.Format("Convert(Value, 'System.String') LIKE '{0}%'", searchTerm)
        '    Else
        '        dv.RowFilter = "" ' Clear the filter if search term is non-numeric
        '    End If
        'End If
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Me.Close()
    End Sub

    Private Sub Form38_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet


        Dim excelWindow = excelApp.ActiveWindow

        Dim cell As Range = worksheet.Range(TargetVar2).Offset(1, 1)
        Dim zoomFactor = excelWindow.Zoom / 100
        'var ws = cell.Worksheet;

        Dim ap = excelWindow.ActivePane ' might be split panes
        Dim origScrollCol = ap.ScrollColumn
        Dim origScrollRow = ap.ScrollRow
        excelApp.ScreenUpdating = False
        ' when FreezePanes == true, ap.ScrollColumn/Row will only reset
        ' as much as the location of the frozen splitter
        ap.ScrollColumn = 1
        ap.ScrollRow = 1

        ' PointsToScreenPixels returns different values if the scroll Is Not currently 1
        ' Temporarily set the scroll back to 1 so that PointsToScreenPixels returns a
        ' value we know how to handle.
        ' (x,y) are screen coordinates for the top left corner of the top left cell
        Dim x As Integer = ap.PointsToScreenPixelsX(0) ' e.g. window.x + row header width
        Dim y As Integer = ap.PointsToScreenPixelsY(0) ' e.g. window.y + ribbon height + column headers height

        Dim dpiX As Single = 0
        Dim dpiY As Single = 0

        Using g As Graphics = Graphics.FromHwnd(IntPtr.Zero)
            dpiX = g.DpiX
            dpiY = g.DpiY
        End Using

        Dim deltaRow As Integer = 0
        Dim deltaCol As Integer = 0
        Dim fromCol As Integer = origScrollCol
        Dim fromRow As Integer = origScrollRow
        If (excelWindow.FreezePanes) Then
            fromCol = 1
            fromRow = 1
            deltaCol = origScrollCol - ap.ScrollColumn '// Note: ap.ScrollColumn/ Row <> 1
            deltaRow = origScrollRow - ap.ScrollRow  '// see comment: when FreezePanes == true ...
        End If

        '// Note Each column width / row height has to be calculated individually.
        '// Before, tried to use this approach:
        '// var r2 = (Microsoft.Office.Interop.Excel.Range) cell.Worksheet.Cells[origScrollRow, origScrollCol];
        '// double dw = cell.Left - r2.Left;
        '// double dh = cell.Top - r2.Top;
        '// However, that only works when the zoom factor Is a whole number.
        '// A fractional zoom (e.g. 1.27) causes each individual row Or column to round to the closest whole number,
        '// which means having to loop through.

        Dim col As Excel.Range
        Dim ww As Double
        Dim newW As Double
        Dim i As Integer
        For i = fromCol To cell.Column - 1
            ' skip the columns between the frozen split and the first visible column
            If i >= ap.ScrollColumn AndAlso i < ap.ScrollColumn + deltaCol Then
                Continue For
            End If

            col = CType(worksheet.Cells(cell.Row, i), Microsoft.Office.Interop.Excel.Range)
            ww = col.Width * dpiX / 72
            newW = zoomFactor * ww
            x += CInt(Math.Round(newW))
        Next


        Dim row As Range
        Dim hh As Double
        Dim newH As Double

        For i = fromRow To cell.Row - 1
            ' skip the rows between the frozen split and the first visible row
            If i >= ap.ScrollRow AndAlso i < ap.ScrollRow + deltaRow Then
                Continue For
            End If

            row = CType(worksheet.Cells(i, cell.Column), Excel.Range)
            hh = row.Height * dpiY / 72
            newH = zoomFactor * hh
            y += CInt(Math.Round(newH))
        Next

        ap.ScrollColumn = origScrollCol
        ap.ScrollRow = origScrollRow
        excelApp.ScreenUpdating = True


        Me.Location = New Point(x, y) + New Point(2, 2)
        ' MsgBox(Me.Location.ToString)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged


        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
        ' Ensure it's not the header row
        If e.RowIndex < 0 Then Return

        'Dim cell As DataGridViewCell = DataGridView1.Rows(e.RowIndex).Cells("Value")


        Dim isChecked As Boolean = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Dim itemValue As String = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString()
        '    MsgBox(1)
        '    If isChecked Then
        'If itemValue = "Select all" And isChecked = True Then


        '    Dim i As Integer = 0
        '    ' Dim j As Integer = 1

        '    For Each r As DataGridViewRow In DataGridView1.Rows

        '        Dim cellValue As String = r.Cells(1).Value.ToString()

        '        'DataGridView1.Rows(i).Cells("Increment").Value = True
        '        r.Cells(0).Value = True
        '        'DataGridView1.Rows(j).Cells("Increment") = True
        '        'MsgBox(1)

        '        i = i + 1
        '    Next




        '    ' DataGridView1.Columns("Increment").cells(0).value = True
        '    'DataGridView1.Rows(e.RowIndex).Cells("Increment").Value = True

        'Else

        '    Dim i As Integer = 0
        '    ' Dim j As Integer = 1

        '    For Each r As DataGridViewRow In DataGridView1.Rows

        '        'Dim cellValue As String = r.Cells(1).Value.ToString()

        '        r.Cells(0).Value = False
        '        'MsgBox(2)
        '        'DataGridView1.Rows(j).Cells("Increment") = True

        '        i = i + 1
        '    Next
        '    Me.Refresh()
        'End If

        If isChecked = True And itemValue <> "Select all" Then
            Me.Refresh()
            ' Place the item in B1 cell
            'worksheet.Range("B1").Value = DataGridView1.Rows(e.RowIndex).Cells("YourItemColumnName").Value
            If worksheet.Range(TargetVar2).Value Is Nothing Then
                worksheet.Range(TargetVar2).Value = itemValue
            Else
                Dim values As String
                Try

                    'values = worksheet.Range(TargetVar2).Value.Split(","c)
                    'values = worksheet.Range(TargetVar2).Value.ToString.Split(New String Separator2, StringSplitOptions.None)
                    'values = worksheet.Range(TargetVar2).Value.ToString().Split(String() Separator2, StringSplitOptions.None)

                    ' Split the string using the separator
                    Dim result() As String = worksheet.Range(TargetVar2).Value.ToString().Split(New String() {Separator2}, StringSplitOptions.None)

                    ' Join the split results back into a single string, using a space as the new separator (or any other separator of your choice)
                    values = String.Join(" ", result)
                Catch ex As Exception
                    values = worksheet.Range(TargetVar2).Value
                End Try
                If values.Contains(itemValue) = False And Horizontal2 = True Then

                    'If Horizontal = True Then
                    worksheet.Range(TargetVar2).Value = worksheet.Range(TargetVar2).Value & Separator2 & itemValue
                ElseIf values.Contains(itemValue) = False And Horizontal2 = False Then
                    worksheet.Range(TargetVar2).Value = worksheet.Range(TargetVar2).Value & Separator2 & vbNewLine & itemValue
                End If
            End If
            'ElseIf e.ColumnIndex = DataGridView1.Columns("Decrement").Index AndAlso e.RowIndex >= 0 Then
            '    Me.Refresh()
            '    Dim itemToRemove As String = cell.Value

            '    If worksheet.Range(TargetVar).Value IsNot Nothing AndAlso worksheet.Range(TargetVar).Value.ToString().Contains(itemToRemove) Then
            '        Dim items As List(Of String) = worksheet.Range(TargetVar).Value.ToString().Split(Separator).ToList()
            '        items.RemoveAll(Function(x) x.Trim() = itemToRemove)
            '        worksheet.Range(TargetVar).Value = String.Join(Separator, items)
            '    End If

        ElseIf isChecked = False Then
            'Me.Refresh()
            Dim itemToRemove As String = itemValue

            If worksheet.Range(TargetVar2).Value IsNot Nothing AndAlso worksheet.Range(TargetVar2).Value.ToString().Contains(itemToRemove) Then
                Dim items As List(Of String) = worksheet.Range(TargetVar2).Value.ToString().Split(New String() {Separator2}, StringSplitOptions.None).ToList()

                ' Find the index of the first occurrence of the item to remove
                Dim indexToRemove As Integer = items.FindIndex(Function(x) x.Trim() = itemToRemove)

                If indexToRemove >= 0 Then ' If found
                    items.RemoveAt(indexToRemove) ' Remove only the first occurrence
                    worksheet.Range(TargetVar2).Value = String.Join(Separator2, items)
                End If
            End If

        End If


        'If e.ColumnIndex = 0 Then
        '    Dim isChecked As Boolean = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        '    Dim itemValue As String = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString()
        '    MsgBox(1)
        '    If isChecked Then
        '        AddToExcelDropdownList(itemValue)
        '        MsgBox(2)
        '    Else
        '        RemoveFromExcelDropdownList(itemValue)
        '        MsgBox(3)
        '    End If
        'End If
    End Sub

    Private Sub AddToExcelDropdownList(value As String)
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
        Dim cell As Excel.Range
        'Dim existingValues As String

        'workbook = excelApp.Workbooks.Open("YOUR_EXCEL_PATH_HERE.xlsx")
        'worksheet = workbook.Worksheets(1)
        MsgBox(4)
        cell = excelApp.Cells(1, 2) ' Change this to your dropdown cell location

        ' Assuming the cell has a dropdown list validation
        'existingValues = cell.Validation.Formula1

        'If Not existingValues.Contains(value) Then
        worksheet.Range(TargetVar2).Value = worksheet.Range(TargetVar2).Value & "," & value

        ' MsgBox(5)

    End Sub

    Private Sub RemoveFromExcelDropdownList(value As String)
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
        Dim cell As Excel.Range

        cell = excelApp.Cells(1, 2) ' Change this to your dropdown cell location

        Dim itemToRemove As String = cell.Value

        If worksheet.Range(TargetVar2).Value IsNot Nothing AndAlso worksheet.Range(TargetVar2).Value.ToString().Contains(itemToRemove) Then
            Dim items As List(Of String) = worksheet.Range(TargetVar2).Value.ToString().Split(New String() {Separator2}, StringSplitOptions.None).ToList()

            ' Find the index of the first occurrence of the item to remove
            Dim indexToRemove As Integer = items.FindIndex(Function(x) x.Trim() = itemToRemove)

            If indexToRemove >= 0 Then ' If found
                items.RemoveAt(indexToRemove) ' Remove only the first occurrence
                worksheet.Range(TargetVar2).Value = String.Join(Separator2, items)
            End If
        End If

    End Sub

    Private Sub Form38_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        'Me.BringToFront()
        Me.Focus()
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        settingflag2 = True
        Me.Hide()
        form = New Form37_MSDropDownCheckBox
        form.Show()
        form.CustomGroupBox6.Enabled = False
        If form Is Nothing Or form.IsDisposed = True Then
            Me.Show()
        End If
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class