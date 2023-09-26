Imports System.Data
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel
Imports DataTable = System.Data.DataTable
Imports Point = System.Drawing.Point

Public Class Form36
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
    Dim power As Boolean

    Private processingEvent As Boolean = False
    Public focuschange As Boolean
    Private Form As Form35Multi_SelectionbasedDropdown = Nothing

    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1

    Public Target As Excel.Range


    Private Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowRect(ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function


    Private Sub Form36_Load(sender As Object, e As EventArgs) Handles Me.Load
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet


        'Enable & Disable Search Option
        If Search1 = True Then
            txtSearch.Enabled = True
            PB_Search.Enabled = True
        Else
            txtSearch.Enabled = False
            PB_Search.Enabled = False
        End If


        ' Add Increment Button
        Dim incrementColumn As New DataGridViewButtonColumn()
        incrementColumn.Name = "Increment"
        incrementColumn.HeaderText = "Add"
        incrementColumn.Text = "+"
        incrementColumn.UseColumnTextForButtonValue = True
        incrementColumn.Width = 28 ' Set the width you want here
        incrementColumn.DefaultCellStyle.BackColor = Color.White
        incrementColumn.FlatStyle = FlatStyle.Popup
        'MsgBox(incrementColumn.DefaultCellStyle.BackColor.ToString)
        incrementColumn.DefaultCellStyle.Font = New System.Drawing.Font("Segoe UI", 12)
        DataGridView1.Columns.Add(incrementColumn)

        ' Add Decrement Button
        Dim decrementColumn As New DataGridViewButtonColumn()
        decrementColumn.Name = "Decrement"
        decrementColumn.HeaderText = "Sub"
        decrementColumn.Text = "-"
        decrementColumn.UseColumnTextForButtonValue = True
        decrementColumn.Width = 28 ' Set the width you want here
        decrementColumn.DefaultCellStyle.BackColor = Color.White
        decrementColumn.FlatStyle = FlatStyle.Popup
        decrementColumn.DefaultCellStyle.Font = New System.Drawing.Font("Segoe UI", 12)
        DataGridView1.Columns.Add(decrementColumn)



        ' Dim formula1 As String
        'Dim dropdownItems() As String
        Dim sourceRange As Excel.Range = Nothing

        ' Get the cell with the drop-down list
        Dim cell As Excel.Range = worksheet.Range(TargetVar1)


        ' Extract the formula (assuming it's a list)
        'formula1 = cell.Validation.Formula1
        'dropdownItems = formula1.Split(","c)


        dt = New DataTable()
        Dim validationList As String = ""
        Dim formula As String = cell.Validation.Formula1

        If formula.Contains(",") Then
            ' Data validation type: Excel Range
        validationList = formula

            dt.Columns.Add("Value", GetType(String))

            Dim items As String() = validationList.Split(","c)
            'dt.Rows(1).Add("Select All")
            For Each item As String In items
                dt.Rows.Add(item)
                'dt.Rows.Add(20)
            Next
            'dt.Rows.Add(20)
        Else
            ' Data validation type: Excel Range

            sourceRange = worksheet.Range(formula)

            dt.Columns.Add("Value", GetType(String))

            'dt.Rows.Add("Select All")
            'Sample Data
            For Each itemCell As Excel.Range In sourceRange
                dt.Rows.Add(itemCell.Value)
                'dt.Rows.Add(20)
                'dt.Rows.Add(30)
            Next
            'dt.Rows.Add(20)

        End If









        '' Extract the formula (assuming it's a range reference)
        'formula1 = cell.Validation.Formula1
        '' Assuming the range is in the same sheet
        'sourceRange = worksheet.Range(formula1)

        ' Populate DataGridView
        'For Each item As String In dropdownItems
        '    DataGridView1.Rows.Add(item)
        'Next


        'DataGridView1.Rows.Add("Select All")
        DataGridView1.DataSource = dt
        DataGridView1.Columns(2).Width = 110


        Dim labelColumn As New DataGridViewTextBoxColumn
        labelColumn.Name = "OccurrenceCount"
        labelColumn.HeaderText = "Occurrences"
        DataGridView1.Columns.Add(labelColumn)
        labelColumn.Width = 72
        labelColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        'Dim excelApp As New Microsoft.Office.Interop.Excel.Application
        'Dim ws As Microsoft.Office.Interop.Excel.Worksheet = excelApp.ActiveSheet
        'Dim targetRange As Excel.Range = ws.Range(rangeAddress)
        Dim targetCellValue As String = worksheet.Range(TargetVar1).Value 'Assuming B1 cell contains the values

        For Each r As DataGridViewRow In DataGridView1.Rows
            Dim itemValue As String = r.Cells(2).Value.ToString()
            'MsgBox(itemValue)
            Dim count As Integer = CountOccurrencesInExcelCell(itemValue, targetCellValue)

            r.Cells("OccurrenceCount").Value = count
        Next

        'DataGridView1.DataSource = dv

        Me.BringToFront()
        Me.Focus()


    End Sub

    Private Function CountOccurrencesInExcelCell(item As String, cellValue As String) As Integer
        'If String.IsNullOrEmpty(cellValue) Then Return 0

        '' Split the cellValue using a comma and remove any extra white spaces
        'Dim items As String() = cellValue.Split(New Char() {","c}).Select(Function(s) s.Trim()).ToArray()

        '' Count occurrences of the item in the split array
        'Return items.Count(Function(i) i = item)
        Dim occurrences As Integer = 0

        If cellValue IsNot Nothing Then
            'MsgBox(1)
            'Dim cellValue As String = targetCell.Value.ToString()
            Dim items As String() = cellValue.Split(New Char() {Separator1}, StringSplitOptions.RemoveEmptyEntries)

            For Each i As String In items
                If i.Trim().Equals(item.Trim(), StringComparison.OrdinalIgnoreCase) Then
                    occurrences += 1
                End If
            Next
        End If
        ' Me.Refresh()
        Return occurrences
    End Function


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


        If e.ColumnIndex = DataGridView1.Columns("Increment").Index Then
            Me.Refresh()
            ' Place the item in B1 cell
            'worksheet.Range("B1").Value = DataGridView1.Rows(e.RowIndex).Cells("YourItemColumnName").Value
            If worksheet.Range(TargetVar1).Value Is Nothing Then
                worksheet.Range(TargetVar1).Value = cell.Value
            Else
                If Horizontal1 = True Then
                    worksheet.Range(TargetVar1).Value = worksheet.Range(TargetVar1).Value & Separator1 & cell.Value
                Else
                    worksheet.Range(TargetVar1).Value = worksheet.Range(TargetVar1).Value & Separator1 & vbNewLine & cell.Value
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

        ElseIf e.ColumnIndex = DataGridView1.Columns("Decrement").Index AndAlso e.RowIndex >= 0 Then
            'Me.Refresh()
            Dim itemToRemove As String = cell.Value

            If worksheet.Range(TargetVar1).Value IsNot Nothing AndAlso worksheet.Range(TargetVar1).Value.ToString().Contains(itemToRemove) Then
                Dim items As List(Of String) = worksheet.Range(TargetVar1).Value.ToString().Split(New String() {Separator1}, StringSplitOptions.None).ToList()

                ' Find the index of the first occurrence of the item to remove
                Dim indexToRemove As Integer = items.FindIndex(Function(x) x.Trim() = itemToRemove)

                If indexToRemove >= 0 Then ' If found
                    items.RemoveAt(indexToRemove) ' Remove only the first occurrence
                    worksheet.Range(TargetVar1).Value = String.Join(Separator1, items)
                End If
            End If

        End If


        'For Each r As DataGridViewRow In DataGridView1.Rows
        '    If txtSearch Is Nothing Then
        '        Dim itemValue As String = r.Cells(2).Value.ToString()
        '        MsgBox(itemValue)
        '        Dim count As Integer = CountOccurrencesInExcelCell(itemValue, worksheet.Range(TargetVar).Value)

        '        r.Cells("OccurrenceCount").Value = count

        '    End If
        'Next



        'If txtSearch.Text = "" Then
        For Each r As DataGridViewRow In DataGridView1.Rows
            'MsgBox(DataGridView1.Rows.Count)
            Dim itemValue As String = r.Cells(2).Value.ToString()
            If power = True Then
                itemValue = r.Cells(3).Value.ToString()
            End If
            'MsgBox(itemValue)
            Dim count As Integer = CountOccurrencesInExcelCell(itemValue, worksheet.Range(TargetVar1).Value)

            r.Cells("OccurrenceCount").Value = count
        Next
        'Else
        'For Each r As DataGridViewRow In dt.Rows
        '        Dim itemValue As String = r.Cells(2).Value.ToString()
        '        'MsgBox(itemValue)
        '        Dim count As Integer = CountOccurrencesInExcelCell(itemValue, worksheet.Range(TargetVar).Value)

        '        r.Cells("OccurrenceCount").Value = count
        '    Next

        'End If

        Me.Refresh()
    End Sub

    'Private Sub btnSearch_Click(sender As Object, e As EventArgs)
    '    Dim searchTerm As String = txtSearch.Text.Trim()
    '    If String.IsNullOrEmpty(searchTerm) Then
    '        DataGridView1.DataSource = dt
    '    Else
    '        Dim dv As New DataView(dt)

    '        If IsNumeric(searchTerm) Then
    '            dv.RowFilter = String.Format("Value = {0}", Convert.ToInt32(searchTerm))
    '            DataGridView1.DataSource = dv
    '        Else
    '            'MessageBox.Show("Please enter a valid number.")
    '        End If
    '    End If
    'End Sub

    'dv.RowFilter = String.Format("Convert(Value, 'System.String') LIKE '{0}%'", searchTerm)

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet

        If txtSearch.Text = "" Or txtSearch.Text = Nothing Then
            power = False
        Else
            power = True
        End If


        Dim searchTerm As String = txtSearch.Text.Trim()
        Dim dv As New DataView(dt)
        If String.IsNullOrEmpty(searchTerm) Then
            DataGridView1.DataSource = dt
        Else
            'Dim dv As New DataView(dt)

            If IsNumeric(searchTerm) Then
                dv.RowFilter = String.Format("Convert(Value, 'System.String') LIKE '{0}%'", searchTerm)

                DataGridView1.DataSource = dv
            Else
                DataGridView1.DataSource = dt
            End If
        End If
        Me.Refresh()
        'MsgBox(DataGridView1.Rows(1).Cells(2).Value.ToString)

        For Each r As DataGridViewRow In DataGridView1.Rows
            Me.Refresh()
            'MsgBox(r.Cells(3).Value)
            Dim itemValue As String = r.Cells(3).Value.ToString
            'MsgBox(itemValue)
            Dim count As Integer = CountOccurrencesInExcelCell(itemValue, worksheet.Range(TargetVar1).Value)

            r.Cells("OccurrenceCount").Value = count
        Next

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
        Me.Refresh()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Me.Close()
    End Sub

    Private Sub Form36_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet


        Dim excelWindow = excelApp.ActiveWindow

        Dim cell As Range = worksheet.Range(TargetVar1).Offset(1, 1)
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

        'myFormInstance = New Form36()
        'yourFormInstance.Show()

        'Form f = New Form();
        'Me.Show()
        'Me.StartPosition = FormStartPosition.Manual
        Me.BringToFront()
        Me.Focus()
        Me.Activate()
        Me.Location = New Point(x, y) + New Point(2, 2)
        ' MsgBox(Me.Location.ToString)


        Me.TopMost = True 'Then it will bring the form to top
        Me.TopMost = False

    End Sub


    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        settingflag1 = True
        Me.Hide()
        Form = New Form35Multi_SelectionbasedDropdown
        Form.Show()
        Form.CustomGroupBox6.Enabled = False
        If Form Is Nothing Or Form.IsDisposed = True Then
            Me.Show()
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Form36_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        'Me.BringToFront()
        Me.Focus()
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class
